import tempfile
import unittest
from pathlib import Path

from search_core import SearchError, column_name, matches, subaddress, value_text, windows_xlsx_path, replace_output


class SearchRulesTests(unittest.TestCase):
    def test_partial_case_sensitive_without_trimming_or_wildcards(self):
        for value, word, expected in [
            ('ABC', 'ABC', True), ('abc', 'ABC', False),
            ('ABC123', 'ABC', True), (' ABC', 'ABC', True),
            ('ABC\n', 'ABC', True), ('*?~', '*?~', True),
            ('def/abc', 'abc', True), ('abc/def', 'abc', True),
            ('def/abc/ghi', 'abc', True), ('abc/abc', 'abc', True),
            ('def/ABC', 'abc', False), ('ab', 'abc', False),
            ('', 'abc', False), ('abc', ' abc', False),
            ('prefix*?~suffix', '*?~', True), ('前一致後', '一致', True),
            ('anything', '*', False), ('一致', '一致', True),
            ('ＡＢＣ', 'ABC', False), (None, 'None', False),
            (123.0, '123', True), (0, '0', True), (True, 'True', True),
        ]:
            with self.subTest(value=value, word=word):
                self.assertEqual(matches(value, word), expected)

    def test_empty_keyword_is_rejected(self):
        for word in ('', None):
            with self.subTest(word=word), self.assertRaisesRegex(SearchError, 'B2'):
                matches('ABC', word)

    def test_cell_address_boundaries_and_quoted_sheet_name(self):
        for number, name in [(1, 'A'), (26, 'Z'), (27, 'AA'), (16384, 'XFD')]:
            self.assertEqual(column_name(number), name)
        self.assertEqual(subaddress("日本語 O'Brien", 1048576, 16384), "'日本語 O''Brien'!XFD1048576")
        for number in (0, 16385):
            with self.assertRaises(ValueError):
                column_name(number)

    def test_numbers_use_values_without_display_formatting(self):
        self.assertEqual(value_text(123.0), '123')
        self.assertEqual(value_text(1.25), '1.25')
        self.assertEqual(value_text(False), 'False')
        self.assertEqual(value_text(None), '')

    def test_full_windows_paths_and_extension(self):
        self.assertEqual(windows_xlsx_path(r'C:\日本語 folder\a.XLSX', 'B1'), r'C:\日本語 folder\a.XLSX')
        self.assertEqual(windows_xlsx_path(r'\\server\share\a.xlsx', 'B1'), r'\\server\share\a.xlsx')
        for path in ('', 'a.xlsx', r'C:relative.xlsx', r'\folder\a.xlsx', 'https://host/a.xlsx', r'C:\a.csv'):
            with self.subTest(path=path), self.assertRaises(SearchError):
                windows_xlsx_path(path, 'B1')

    def test_replace_existing_result_and_failed_replace_preserves_it(self):
        with tempfile.TemporaryDirectory() as directory:
            dest = Path(directory) / 'result.xlsx'
            temp = Path(directory) / 'temporary.xlsx'
            dest.write_bytes(b'previous output')
            temp.write_bytes(b'complete new output')
            replace_output(temp, dest)
            self.assertEqual(dest.read_bytes(), b'complete new output')
            self.assertFalse(temp.exists())
            with self.assertRaises(FileNotFoundError):
                replace_output(temp, dest)
            self.assertEqual(dest.read_bytes(), b'complete new output')


if __name__ == '__main__':
    unittest.main(verbosity=2)
