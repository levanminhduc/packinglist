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


class TestClearQuantityColumnsBulk(unittest.TestCase):

    def setUp(self):
        with patch.object(ExcelCOMManager, '__init__', lambda self, *a, **kw: None):
            self.manager = ExcelCOMManager()
            self.manager.config = MagicMock()
            self.manager.config.get_start_row.return_value = 19
            self.manager.excel_app = MagicMock()
            self.screen_updating_mock = PropertyMock()
            type(self.manager.excel_app).ScreenUpdating = self.screen_updating_mock
            self.manager.worksheet = MagicMock()

    def test_calls_clear_contents_on_range(self):
        self.manager.clear_quantity_columns(start_row=19, end_row=59, start_col=7, end_col=39)

        self.manager.worksheet.Range.assert_called_once_with("G19:AM59")
        self.manager.worksheet.Range.return_value.ClearContents.assert_called_once()

    def test_returns_count_from_count_a(self):
        self.manager.excel_app.WorksheetFunction.CountA.return_value = 42

        result = self.manager.clear_quantity_columns(start_row=19, end_row=59, start_col=7, end_col=39)

        self.assertEqual(result, 42)

    def test_screen_updating_toggled(self):
        self.manager.clear_quantity_columns(start_row=19, end_row=59, start_col=7, end_col=39)

        set_calls = [args[0] for args, kwargs in self.screen_updating_mock.call_args_list]

        self.assertIn(False, set_calls)

    def test_screen_updating_restored_on_error(self):
        self.manager.worksheet.Range.return_value.ClearContents.side_effect = Exception("COM error")

        with self.assertRaises(RuntimeError):
            self.manager.clear_quantity_columns(start_row=19, end_row=59, start_col=7, end_col=39)

        set_calls = [args[0] for args, kwargs in self.screen_updating_mock.call_args_list]
        self.assertIn(True, set_calls)

    def test_raises_if_no_worksheet(self):
        self.manager.worksheet = None
        with self.assertRaises(RuntimeError):
            self.manager.clear_quantity_columns()


if __name__ == "__main__":
    unittest.main()
