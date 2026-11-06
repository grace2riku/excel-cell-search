"""Integration tests against real Excel, no mocks. Requires Windows."""
import gc
import os
from pathlib import Path
import sys
import tempfile
import unittest
from urllib.parse import unquote, urlparse

from excel_search import search
from search_core import SearchError


@unittest.skipUnless(sys.platform == 'win32', 'Windows desktop Excel is required')
class ExcelIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import pythoncom
        import win32com.client
        cls.pythoncom = pythoncom
        # pythoncom owns the main thread's COM initialization for this process.
        cls.excel = None
        cls.folder = Path(tempfile.mkdtemp(prefix='ExcelSearchPythonTest_'))
        print(f'\nExcel test artifacts: {cls.folder}')
        cls.addClassCleanup(cls.cleanup_excel)
        cls.excel = win32com.client.DispatchEx('Excel.Application')
        cls.excel.Visible = False
        cls.excel.DisplayAlerts = False
        cls.excel.AutomationSecurity = 3
        cls.source = cls.folder / '日本語 source.xlsx'
        book = cls.excel.Workbooks.Add(-4167)
        try:
            sheet = book.Worksheets(1)
            sheet.Name = "日本語 O'Brien"
            sheet.Range('A1:A8').Value2 = (('ABC',), ('abc',), ('ABC123',),
                                           ('unused',), (' ABC',), ('*?~',), ('anything',), (123,))
            sheet.Cells(4, 1).Formula = '="AB"&"C"'
            hidden = book.Worksheets.Add()
            hidden.Name = 'Hidden'
            hidden.Cells(2, 2).Value2 = 'ABC'
            hidden.Visible = 0
            book.SaveAs(str(cls.source), 51)
        finally:
            book.Close(False)

    @classmethod
    def cleanup_excel(cls):
        try:
            if cls.excel is not None:
                cls.excel.Quit()
        finally:
            cls.excel = None
            gc.collect()

    def setUp(self):
        self.config = self.folder / (self._testMethodName + '_config.xlsx')
        self.output = self.folder / (self._testMethodName + '_result.xlsx')

    def configure(self, word='ABC', output=None, source=None):
        book = self.excel.Workbooks.Add(-4167)
        try:
            sheet = book.Worksheets(1)
            sheet.Range('B1:B3').Value2 = ((str(source or self.source),), (word,), (str(output or self.output),))
            book.SaveAs(str(self.config), 51)
        finally:
            book.Close(False)

    def test_all_sheets_partial_case_sensitive_formula_results_and_real_link_targets(self):
        self.configure()
        count, output = search(self.config)
        self.assertEqual(count, 5)
        self.assertEqual(output, self.output)
        book = self.excel.Workbooks.Open(str(output), 0, True)
        try:
            sheet = book.Worksheets(1)
            self.assertEqual(sheet.Range('A1:B1').Value2, (('No', 'ハイパーリンク'),))
            self.assertEqual(sheet.Hyperlinks.Count, 5)
            targets = []
            for row in range(2, 7):
                self.assertEqual(sheet.Cells(row, 1).Value2, row - 1)
                link = sheet.Cells(row, 2).Hyperlinks(1)
                address = link.Address
                if address.lower().startswith('file:'):
                    parsed = urlparse(address)
                    address = unquote(parsed.path).lstrip('/') if not parsed.netloc else '\\\\' + parsed.netloc + unquote(parsed.path)
                path = Path(address)
                if not path.is_absolute():
                    path = self.source.parent / path
                self.assertEqual(os.path.normcase(str(path.resolve())), os.path.normcase(str(self.source.resolve())))
                targets.append(link.SubAddress)
            self.assertCountEqual(targets, ["'Hidden'!B2", "'日本語 O''Brien'!A1", "'日本語 O''Brien'!A3",
                                            "'日本語 O''Brien'!A4", "'日本語 O''Brien'!A5"])
        finally:
            book.Close(False)

    def test_search_preserves_existing_excel_connection_on_success_and_failure(self):
        # Regression: search used to uninitialize the caller's COM apartment,
        # disconnecting this independent Excel instance after the first call.
        sentinel = self.excel.Workbooks.Add(-4167)
        try:
            sentinel.Worksheets(1).Cells(1, 1).Value2 = '接続維持の確認'
            for _ in range(2):
                self.configure('ABC')
                self.assertEqual(search(self.config)[0], 5)
                self.assertEqual(sentinel.Worksheets(1).Cells(1, 1).Value2, '接続維持の確認')
            self.configure('')
            with self.assertRaisesRegex(SearchError, 'B2'):
                search(self.config)
            self.assertEqual(sentinel.Worksheets(1).Cells(1, 1).Value2, '接続維持の確認')
            self.configure('ABC')
            self.assertEqual(search(self.config)[0], 5)
            self.assertEqual(sentinel.Worksheets(1).Cells(1, 1).Value2, '接続維持の確認')
        finally:
            sentinel.Close(False)

    def test_literal_wildcards_space_numeric_and_zero_overwrites_previous_results(self):
        for word, expected in [('*?~', 1), (' ABC', 1), ('123', 2), ('not found', 0)]:
            with self.subTest(word=word):
                self.configure(word)
                self.assertEqual(search(self.config)[0], expected)
        book = self.excel.Workbooks.Open(str(self.output), 0, True)
        try:
            self.assertEqual(book.Worksheets(1).Hyperlinks.Count, 0)
            self.assertIsNone(book.Worksheets(1).Cells(2, 1).Value2)
        finally:
            book.Close(False)

    def test_invalid_configuration_preserves_existing_result(self):
        self.output.write_bytes(b'previous output must survive')
        for word, dest, source, message in [
            ('', self.output, self.source, 'B2'),
            ('ABC', self.source, self.source, '同じ'),
            ('ABC', self.config, self.source, '同じ'),
            ('ABC', self.output, self.folder / 'missing.xlsx', '見つかりません'),
            ('ABC', self.folder / 'missing' / 'result.xlsx', self.source, 'フォルダー'),
            ('ABC', self.folder / 'result.csv', self.source, '.xlsx'),
        ]:
            with self.subTest(message=message):
                self.configure(word, dest, source)
                with self.assertRaisesRegex(SearchError, message):
                    search(self.config)
                self.assertEqual(self.output.read_bytes(), b'previous output must survive')

    def test_locked_destination_is_not_replaced(self):
        self.configure()
        self.output.write_bytes(b'locked existing output')
        with self.output.open('rb'):
            with self.assertRaises(OSError):
                search(self.config)
        self.assertEqual(self.output.read_bytes(), b'locked existing output')

    def test_one_cell_usedrange(self):
        source = self.folder / 'single.xlsx'
        book = self.excel.Workbooks.Add(-4167)
        try:
            book.Worksheets(1).Cells(1, 1).Value2 = '一致'
            book.SaveAs(str(source), 51)
        finally:
            book.Close(False)
        self.configure('一致', source=source)
        self.assertEqual(search(self.config)[0], 1)

    def test_partial_match_links_once_per_cell_including_formula_result(self):
        source = self.folder / 'partial.xlsx'
        book = self.excel.Workbooks.Add(-4167)
        try:
            sheet = book.Worksheets(1)
            sheet.Name = '部分一致'
            sheet.Range('A1:A5').Value2 = (('def/abc',), ('abc/abc',), ('def/ABC',), ('unused',), ('ab',))
            sheet.Cells(4, 1).Formula = '="def/"&"abc"'
            book.SaveAs(str(source), 51)
        finally:
            book.Close(False)
        self.configure('abc', source=source)
        self.assertEqual(search(self.config)[0], 3)
        book = self.excel.Workbooks.Open(str(self.output), 0, True)
        try:
            sheet = book.Worksheets(1)
            self.assertEqual(sheet.Hyperlinks.Count, 3)
            self.assertEqual([sheet.Cells(r, 1).Value2 for r in range(2, 5)], [1, 2, 3])
            self.assertEqual([sheet.Cells(r, 2).Hyperlinks(1).SubAddress for r in range(2, 5)],
                             ["'部分一致'!A1", "'部分一致'!A2", "'部分一致'!A4"])
        finally:
            book.Close(False)


if __name__ == '__main__':
    unittest.main(verbosity=2)
