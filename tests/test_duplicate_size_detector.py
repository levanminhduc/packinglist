import unittest
from unittest.mock import MagicMock, patch, PropertyMock
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_automation.duplicate_size_detector import DuplicateSizeDetector


class TestDetectDuplicates(unittest.TestCase):

    def setUp(self):
        self.com_manager = MagicMock()
        self.com_manager.worksheet = MagicMock()
        self.com_manager.config = MagicMock()
        self.com_manager.config.get_column.return_value = "F"
        self.com_manager.config.get_start_row.return_value = 19
        self.com_manager.detect_end_row.return_value = 23
        self.detector = DuplicateSizeDetector(self.com_manager)

    def test_no_duplicates_returns_empty(self):
        self.com_manager.worksheet.Range.return_value.Value = (
            (44.0,), (45.0,), (46.0,), (47.0,), (48.0,),
        )
        result = self.detector.detect_duplicates()
        self.assertEqual(result, {})

    def test_detects_duplicate_sizes(self):
        self.com_manager.worksheet.Range.return_value.Value = (
            (60.0,), (44.0,), (60.0,), (45.0,), (44.0,),
        )
        result = self.detector.detect_duplicates()
        self.assertIn("060", result)
        self.assertEqual(result["060"], [19, 21])
        self.assertIn("044", result)
        self.assertEqual(result["044"], [20, 23])
        self.assertNotIn("045", result)

    def test_none_values_skipped(self):
        self.com_manager.worksheet.Range.return_value.Value = (
            (60.0,), (None,), (60.0,), (None,), (None,),
        )
        result = self.detector.detect_duplicates()
        self.assertIn("060", result)
        self.assertEqual(result["060"], [19, 21])

    def test_empty_range_returns_empty(self):
        self.com_manager.worksheet.Range.return_value.Value = None
        result = self.detector.detect_duplicates()
        self.assertEqual(result, {})


class TestDeleteRows(unittest.TestCase):

    def setUp(self):
        self.com_manager = MagicMock()
        self.com_manager.worksheet = MagicMock()
        self.com_manager.excel_app = MagicMock()
        self.screen_updating = PropertyMock()
        type(self.com_manager.excel_app).ScreenUpdating = self.screen_updating
        self.detector = DuplicateSizeDetector(self.com_manager)

    def test_deletes_rows_bottom_up(self):
        rows_to_delete = [19, 25, 31]
        result = self.detector.delete_rows(rows_to_delete)

        self.assertEqual(result, 3)

        calls = self.com_manager.worksheet.Rows.call_args_list
        deleted_rows = [call[0][0] for call in calls]
        self.assertEqual(deleted_rows, [31, 25, 19])

    def test_screen_updating_toggled(self):
        self.detector.delete_rows([19, 25])

        self.screen_updating.assert_any_call(False)
        self.screen_updating.assert_any_call(True)

    def test_empty_list_returns_zero(self):
        result = self.detector.delete_rows([])
        self.assertEqual(result, 0)

    def test_screen_updating_restored_on_error(self):
        self.com_manager.worksheet.Rows.return_value.Delete.side_effect = Exception("COM error")

        with self.assertRaises(RuntimeError):
            self.detector.delete_rows([19])

        self.screen_updating.assert_called_with(True)


if __name__ == '__main__':
    unittest.main()
