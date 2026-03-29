import unittest
from unittest.mock import MagicMock, patch, PropertyMock

from ui.pdf_import_dialog import ImportProgressDialog


class TestImportProgressDialogSteps(unittest.TestCase):

    def test_steps_count_is_9(self):
        assert len(ImportProgressDialog.STEPS) == 9

    def test_step_weights_count_matches_steps(self):
        assert len(ImportProgressDialog.STEP_WEIGHTS) == len(ImportProgressDialog.STEPS)

    def test_step_weights_sum_to_100(self):
        assert sum(ImportProgressDialog.STEP_WEIGHTS) == 100

    def test_new_steps_exist(self):
        assert "Ghi Sizes & Quantities vào Excel" in ImportProgressDialog.STEPS
        assert "Ẩn dòng không chọn" in ImportProgressDialog.STEPS

    def test_new_steps_before_hoan_tat(self):
        idx_write = ImportProgressDialog.STEPS.index("Ghi Sizes & Quantities vào Excel")
        idx_hide = ImportProgressDialog.STEPS.index("Ẩn dòng không chọn")
        idx_done = ImportProgressDialog.STEPS.index("Hoàn tất")
        assert idx_write < idx_hide < idx_done

    def test_hoan_tat_is_last_step(self):
        assert ImportProgressDialog.STEPS[-1] == "Hoàn tất"


if __name__ == "__main__":
    unittest.main()
