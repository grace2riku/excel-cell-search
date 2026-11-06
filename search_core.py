"""Platform-independent search rules and safe output replacement."""
import ntpath
import os


class SearchError(ValueError):
    pass


def value_text(value):
    if value is None:
        return ''
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        return format(value, '.15g')
    return str(value)


def matches(value, keyword):
    if keyword is None or keyword == '':
        raise SearchError('B2の検索ワードが空です。')
    return value is not None and keyword in value_text(value)


def windows_xlsx_path(value, label):
    if not isinstance(value, str) or not value.strip():
        raise SearchError(f'{label}が空です。')
    drive, tail = ntpath.splitdrive(value)
    if not drive or not tail.startswith(('\\', '/')) or (not drive.startswith('\\\\') and not (len(drive) == 2 and drive[1] == ':')):
        raise SearchError(f'{label}にはWindowsのフルパスを指定してください。')
    if ntpath.splitext(value)[1].lower() != '.xlsx':
        raise SearchError(f'{label}は.xlsxファイルにしてください。')
    return ntpath.normpath(value)


def column_name(number):
    if not 1 <= number <= 16384:
        raise ValueError('Excel列番号の範囲外です。')
    result = ''
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def subaddress(sheet, row, column):
    if not 1 <= row <= 1048576:
        raise ValueError('Excel行番号の範囲外です。')
    return "'" + sheet.replace("'", "''") + "'!" + column_name(column) + str(row)


def replace_output(temporary, destination):
    os.replace(temporary, destination)
