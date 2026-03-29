import unittest
from unittest.mock import MagicMock, patch, PropertyMock
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_automation.excel_com_manager import ExcelCOMManager


class TestNumberToColumnLetter(unittest.TestCase):

    def setUp(self):
        with patch.object(ExcelCOMManager, '__init__', lambda self, *a, **kw: None):
            self.manager = ExcelCOMManager()

    def test_single_letter_columns(self):
        self.assertEqual(self.manager._number_to_column_letter(1), "A")
        self.assertEqual(self.manager._number_to_column_letter(7), "G")
        self.assertEqual(self.manager._number_to_column_letter(26), "Z")

    def test_double_letter_columns(self):
        self.assertEqual(self.manager._number_to_column_letter(27), "AA")
        self.assertEqual(self.manager._number_to_column_letter(39), "AM")

    def test_column_letter_roundtrip(self):
        self.manager._column_letter_to_number = ExcelCOMManager._column_letter_to_number.__get__(self.manager)
        for col_num in [1, 7, 26, 27, 39]:
            letter = self.manager._number_to_column_letter(col_num)
            result = self.manager._column_letter_to_number(letter)
            self.assertEqual(result, col_num)


if __name__ == "__main__":
    unittest.main()
