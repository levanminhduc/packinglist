import unittest
from unittest.mock import MagicMock, patch, ANY
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))


class TestBoxListExportProgressDialog(unittest.TestCase):

    def setUp(self):
        self.tk_patcher = patch('ui.box_list_export_progress_dialog.tk')
        self.ttk_patcher = patch('ui.box_list_export_progress_dialog.ttk')
        self.mock_tk = self.tk_patcher.start()
        self.mock_ttk = self.ttk_patcher.start()

        self.mock_parent = MagicMock()
        self.mock_dialog = MagicMock()
        self.mock_tk.Toplevel.return_value = self.mock_dialog
        self.mock_tk.IntVar.return_value = MagicMock()
        self.mock_tk.BOTH = 'both'
        self.mock_tk.W = 'w'
        self.mock_tk.X = 'x'
        self.mock_tk.LEFT = 'left'
        self.mock_tk.RIGHT = 'right'

        from ui.box_list_export_progress_dialog import BoxListExportProgressDialog
        self.dialog = BoxListExportProgressDialog(self.mock_parent)

    def tearDown(self):
        self.tk_patcher.stop()
        self.ttk_patcher.stop()

    def test_has_6_steps(self):
        self.assertEqual(len(self.dialog.STEPS), 6)

    def test_step_weights_sum_to_100(self):
        self.assertEqual(sum(self.dialog.STEP_WEIGHTS), 100)

    def test_step_weights_length_matches_steps(self):
        self.assertEqual(len(self.dialog.STEP_WEIGHTS), len(self.dialog.STEPS))

    def test_dialog_is_modal(self):
        self.mock_dialog.transient.assert_called_once_with(self.mock_parent)
        self.mock_dialog.grab_set.assert_called_once()

    def test_dialog_blocks_close(self):
        self.mock_dialog.protocol.assert_called_with("WM_DELETE_WINDOW", ANY)

    def test_start_step_updates_current_step(self):
        self.dialog.start_step(3)
        self.assertEqual(self.dialog.current_step, 3)

    def test_start_step_updates_progress_percent(self):
        self.dialog.start_step(2)
        expected_percent = sum(self.dialog.STEP_WEIGHTS[:2])
        self.mock_tk.IntVar.return_value.set.assert_called_with(expected_percent)

    def test_complete_step_updates_progress_percent(self):
        self.dialog.complete_step(1)
        expected_percent = sum(self.dialog.STEP_WEIGHTS[:2])
        self.mock_tk.IntVar.return_value.set.assert_called_with(expected_percent)

    def test_finish_sets_100_percent(self):
        self.dialog.finish()
        self.mock_tk.IntVar.return_value.set.assert_called_with(100)

    def test_finish_auto_closes_after_delay(self):
        self.dialog.finish()
        self.mock_parent.after.assert_called_once_with(1000, self.mock_dialog.destroy)

    def test_show_error_stores_retry_callback(self):
        callback = MagicMock()
        self.dialog.show_error(1, "test error", callback)
        self.assertEqual(self.dialog.retry_callback, callback)

    def test_show_error_allows_close(self):
        callback = MagicMock()
        self.dialog.show_error(1, "test error", callback)
        self.mock_dialog.protocol.assert_called_with("WM_DELETE_WINDOW", self.mock_dialog.destroy)

    def test_close_destroys_dialog(self):
        self.dialog.close()
        self.mock_dialog.destroy.assert_called_once()

    def test_retry_invokes_callback(self):
        callback = MagicMock()
        self.dialog.show_error(1, "test error", callback)
        self.dialog._retry()
        callback.assert_called_once()

    def test_steps_contain_expected_names(self):
        step_names = self.dialog.STEPS
        self.assertIn("Đọc dữ liệu thùng từ Excel", step_names)
        self.assertIn("Phân tích & gộp sizes", step_names)
        self.assertIn("Tạo sheet mới", step_names)
        self.assertIn("Ghi danh sách thùng vào sheet", step_names)
        self.assertIn("Copy vào clipboard", step_names)
        self.assertIn("Hoàn tất", step_names)


if __name__ == "__main__":
    unittest.main()
