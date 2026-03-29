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

    def test_uses_detected_tot_qty_column_as_end(self):
        with patch.object(self.manager, '_detect_tot_qty_column', return_value=14):
            self.manager.clear_quantity_columns(start_row=19, end_row=59)

        self.manager.worksheet.Range.assert_called_once_with("G19:M59")
        self.manager.worksheet.Range.return_value.ClearContents.assert_called_once()

    def test_fallback_to_39_when_tot_qty_not_found(self):
        with patch.object(self.manager, '_detect_tot_qty_column', return_value=None):
            self.manager.clear_quantity_columns(start_row=19, end_row=59)

        self.manager.worksheet.Range.assert_called_once_with("G19:AM59")

    def test_explicit_end_col_skips_detect(self):
        with patch.object(self.manager, '_detect_tot_qty_column') as mock_detect:
            self.manager.clear_quantity_columns(start_row=19, end_row=59, end_col=25)

        mock_detect.assert_not_called()
        self.manager.worksheet.Range.assert_called_once_with("G19:Y59")


class TestShowAllRowsBulk(unittest.TestCase):

    def setUp(self):
        with patch.object(ExcelCOMManager, '__init__', lambda self, *a, **kw: None):
            self.manager = ExcelCOMManager()
            self.manager.config = MagicMock()
            self.manager.config.get_start_row.return_value = 19
            self.manager.config.get_end_row.return_value = 59
            self.manager.excel_app = MagicMock()
            self.manager.worksheet = MagicMock()

    def test_unhides_entire_range(self):
        self.manager.show_all_rows(start_row=19, end_row=59)

        self.manager.worksheet.Range.assert_called_once_with("19:59")

    def test_raises_if_no_worksheet(self):
        self.manager.worksheet = None
        with self.assertRaises(RuntimeError):
            self.manager.show_all_rows()

    def test_screen_updating_restored_on_error(self):
        self.manager.worksheet.Range.side_effect = Exception("COM error")

        with self.assertRaises(RuntimeError):
            self.manager.show_all_rows(start_row=19, end_row=59)


class TestScanSizesBulk(unittest.TestCase):

    def setUp(self):
        with patch.object(ExcelCOMManager, '__init__', lambda self, *a, **kw: None):
            self.manager = ExcelCOMManager()
            self.manager.config = MagicMock()
            self.manager.config.get_column.return_value = "F"
            self.manager.config.get_start_row.return_value = 19
            self.manager.worksheet = MagicMock()

        self.manager._column_letter_to_number = ExcelCOMManager._column_letter_to_number.__get__(self.manager)

    def test_reads_range_in_single_call(self):
        self.manager.worksheet.Range.return_value.Value = (
            ("044",), ("045",), ("046",), (None,), ("044",)
        )

        with patch.object(self.manager, 'detect_end_row', return_value=23):
            result = self.manager.scan_sizes(column="F", start_row=19, end_row=23)

        self.manager.worksheet.Range.assert_called_once_with("F19:F23")

    def test_returns_unique_sorted_sizes(self):
        self.manager.worksheet.Range.return_value.Value = (
            ("046",), ("044",), ("045",), ("044",), ("046",)
        )

        with patch.object(self.manager, 'detect_end_row', return_value=23):
            result = self.manager.scan_sizes(column="F", start_row=19, end_row=23)

        self.assertEqual(result, ["044", "045", "046"])

    def test_handles_float_values(self):
        self.manager.worksheet.Range.return_value.Value = (
            (44.0,), (45.0,), (46.0,)
        )

        with patch.object(self.manager, 'detect_end_row', return_value=21):
            with patch.object(self.manager, '_fix_decimal_cell'):
                result = self.manager.scan_sizes(column="F", start_row=19, end_row=21)

        self.assertIn("044", result)
        self.assertIn("045", result)
        self.assertIn("046", result)

    def test_skips_none_values(self):
        self.manager.worksheet.Range.return_value.Value = (
            (None,), ("044",), (None,)
        )

        with patch.object(self.manager, 'detect_end_row', return_value=21):
            result = self.manager.scan_sizes(column="F", start_row=19, end_row=21)

        self.assertEqual(result, ["044"])

    def test_single_cell_returns_non_tuple(self):
        self.manager.worksheet.Range.return_value.Value = "044"

        with patch.object(self.manager, 'detect_end_row', return_value=19):
            result = self.manager.scan_sizes(column="F", start_row=19, end_row=19)

        self.assertEqual(result, ["044"])

    def test_raises_if_no_worksheet(self):
        self.manager.worksheet = None
        with self.assertRaises(RuntimeError):
            self.manager.scan_sizes()


class TestDetectTotQtyColumn(unittest.TestCase):

    def setUp(self):
        with patch.object(ExcelCOMManager, '__init__', lambda self, *a, **kw: None):
            self.manager = ExcelCOMManager()
            self.manager.worksheet = MagicMock()

    def _make_row_values(self, col_count: int, tot_qty_col: int = None, text: str = "Tot QTY"):
        row = [None] * col_count
        if tot_qty_col is not None:
            row[tot_qty_col - 1] = text
        return tuple(row)

    def test_finds_tot_qty_at_column_14(self):
        row_data = self._make_row_values(52, tot_qty_col=14)
        self.manager.worksheet.Range.return_value.Value = (row_data,)
        result = self.manager._detect_tot_qty_column()
        self.assertEqual(result, 14)

    def test_finds_tot_qty_at_column_40(self):
        row_data = self._make_row_values(52, tot_qty_col=40)
        self.manager.worksheet.Range.return_value.Value = (row_data,)
        result = self.manager._detect_tot_qty_column()
        self.assertEqual(result, 40)

    def test_finds_total_qty_variant(self):
        row_data = self._make_row_values(52, tot_qty_col=20, text="Total QTY")
        self.manager.worksheet.Range.return_value.Value = (row_data,)
        result = self.manager._detect_tot_qty_column()
        self.assertEqual(result, 20)

    def test_case_insensitive(self):
        row_data = self._make_row_values(52, tot_qty_col=14, text="tot qty")
        self.manager.worksheet.Range.return_value.Value = (row_data,)
        result = self.manager._detect_tot_qty_column()
        self.assertEqual(result, 14)

    def test_returns_none_when_not_found(self):
        row_data = self._make_row_values(52)
        self.manager.worksheet.Range.return_value.Value = (row_data,)
        result = self.manager._detect_tot_qty_column()
        self.assertIsNone(result)

    def test_returns_none_when_no_worksheet(self):
        self.manager.worksheet = None
        result = self.manager._detect_tot_qty_column()
        self.assertIsNone(result)


if __name__ == "__main__":
    unittest.main()
