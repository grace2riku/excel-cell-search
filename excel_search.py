"""Windows desktop Excel adapter; Python 3.12 + pywin32."""
import argparse
import gc
import os
from pathlib import Path
import sys
import tempfile
import traceback

from search_core import SearchError, matches, replace_output, subaddress, value_text, windows_xlsx_path

MAX_HYPERLINKS = 65530  # Excel documented per-worksheet limit.


def excel_text(value, cell=None):
    # Excel numeric values are doubles; error VARIANTs are signed integers.
    if type(value) is int and value < 0 and cell is not None:
        return str(cell.Text)
    return value_text(value)


def same_file(left, right):
    return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right)) or (
        os.path.exists(left) and os.path.exists(right) and os.path.samefile(left, right)
    )


def search(input_path, progress=print):
    """Search on the caller's COM-initialized thread.

    pythoncom initializes the main Python thread on import. A custom worker
    thread must manage its own COM lifetime outside this function. Never
    uninitialize the caller's apartment while its other COM objects are live.
    """
    if sys.platform != 'win32':
        raise SearchError('このプログラムにはWindowsとデスクトップ版Excelが必要です。')
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise SearchError('pywin32がありません。先にSetup.cmdを実行してください。') from exc
    input_path = Path(windows_xlsx_path(str(input_path), '入力ファイル'))
    if not input_path.is_file():
        raise SearchError(f'入力ファイルが見つかりません: {input_path}')
    excel = None
    config = source_book = result_book = None
    settings = sheets = sheet = used = block = result = cell = None
    temporary = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3
        # UpdateLinks=0, ReadOnly=True, empty passwords prevent password prompts.
        config = excel.Workbooks.Open(str(input_path), 0, True, pythoncom.Missing, '', '', True)
        excel.CalculateFull()
        settings = config.Worksheets(1)
        source = Path(windows_xlsx_path(value_text(settings.Range('B1').Value2), 'B1'))
        cell = settings.Range('B2')
        keyword = excel_text(cell.Value2, cell)
        matches('', keyword)  # Validate empty keyword using the search rule.
        destination = Path(windows_xlsx_path(value_text(settings.Range('B3').Value2), 'B3'))
        if same_file(destination, input_path) or same_file(destination, source):
            raise SearchError('出力先に入力ファイルまたは検索対象と同じファイルは指定できません。')
        if not source.is_file():
            raise SearchError(f'検索対象が見つかりません: {source}')
        if not destination.parent.is_dir():
            raise SearchError(f'出力フォルダーが存在しません: {destination.parent}')
        config.Close(False)
        config = settings = cell = None
        source_book = excel.Workbooks.Open(str(source), 0, True, pythoncom.Missing, '', '', True)
        excel.CalculateFull()
        result_book = excel.Workbooks.Add(-4167)
        result = result_book.Worksheets(1)
        result.Name = '検索結果'
        result.Range('A1:B1').Value2 = (('No', 'ハイパーリンク'),)
        result_book.BuiltinDocumentProperties('Hyperlink Base').Value = str(source.parent)
        count = 0
        sheets = source_book.Worksheets
        for index in range(1, sheets.Count + 1):
            sheet = sheets(index)
            progress(f'検索中: {sheet.Name} ({index}/{sheets.Count})')
            used = sheet.UsedRange
            row_start, column_start = used.Row, used.Column
            row_count, column_count = used.Rows.Count, used.Columns.Count
            block_height = max(1, 100000 // column_count)
            for offset in range(0, row_count, block_height):
                height = min(block_height, row_count - offset)
                block = sheet.Range(sheet.Cells(row_start + offset, column_start),
                                    sheet.Cells(row_start + offset + height - 1, column_start + column_count - 1))
                values = block.Value2
                if height == column_count == 1:
                    values = ((values,),)
                for r, row_values in enumerate(values):
                    for c, value in enumerate(row_values):
                        if value is None:
                            continue
                        row, column = row_start + offset + r, column_start + c
                        if type(value) is int and value < 0:
                            value = excel_text(value, sheet.Cells(row, column))
                        if not matches(value, keyword):
                            continue
                        if count >= MAX_HYPERLINKS:
                            raise SearchError('ヒット数が65,530件を超えました。既存の結果は更新しません。')
                        target = subaddress(sheet.Name, row, column)
                        label = f'{source.name} / {sheet.Name}!{target.rsplit("!", 1)[1]}'
                        result.Hyperlinks.Add(result.Cells(count + 2, 2), str(source), target,
                                              'クリックして該当セルへ移動', label)
                        count += 1
                        result.Cells(count + 1, 1).Value2 = count
                block = None
            used = sheet = None
        result.Range('A1:B1').Font.Bold = True
        result.Range('A1:B1').Interior.Color = 15132390
        result.Columns(1).ColumnWidth = 10
        result.Columns(2).ColumnWidth = 85
        fd, temporary = tempfile.mkstemp(prefix='.excel_search_', suffix='.xlsx', dir=destination.parent)
        os.close(fd)
        result_book.SaveAs(temporary, 51)
        result_book.Close(False)
        result_book = result = None
        source_book.Close(False)
        source_book = sheets = None
        replace_output(temporary, destination)
        temporary = None
        return count, destination
    finally:
        for book in (result_book, source_book, config):
            if book is not None:
                try:
                    book.Close(False)
                except Exception:
                    pass
        book = None
        settings = sheets = sheet = used = block = result = cell = None
        config = source_book = result_book = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        excel = None
        gc.collect()
        if temporary and os.path.isfile(temporary):
            os.unlink(temporary)


def main():
    parser = argparse.ArgumentParser(description='Excelセル部分一致検索（大文字・小文字を区別）')
    parser.add_argument('input', nargs='?', help='B1～B3を設定した.xlsxのフルパス')
    args = parser.parse_args()
    try:
        path = args.input
        if not path:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            try:
                path = filedialog.askopenfilename(title='B1～B3を設定した入力ファイルを選択',
                                                 filetypes=[('Excel workbook', '*.xlsx')])
            finally:
                root.destroy()
            if not path:
                return 0
        count, destination = search(path)
        print(f'完了: {count}件\n出力先: {destination}')
        return 0
    except Exception as exc:
        print(f'エラー: {exc}', file=sys.stderr)
        log = Path(tempfile.gettempdir()) / 'ExcelCellSearchPython-error.log'
        try:
            log.write_text(traceback.format_exc(), encoding='utf-8')
            print(f'詳細ログ: {log}', file=sys.stderr)
        except OSError:
            traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
