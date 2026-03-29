import unittest
from unittest.mock import MagicMock, patch, PropertyMock
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))


def make_mock_worksheet(size_column_data, box_start_row_data, box_end_row_data, size_row_data_map):
    worksheet = MagicMock()

    used_range = MagicMock()
    used_range.Column = 1
    used_range.Columns.Count = 38
    worksheet.UsedRange = used_range

    def range_side_effect(*args):
        range_obj = MagicMock()

        if len(args) == 1 and isinstance(args[0], str):
            range_obj.Value = size_column_data
            return range_obj

        if len(args) == 2:
            start_cell, end_cell = args
            row = start_cell._row
            if row in size_row_data_map:
                range_obj.Value = size_row_data_map[row]
            elif row == box_start_row_data['row']:
                range_obj.Value = box_start_row_data['values']
            elif row == box_end_row_data['row']:
                range_obj.Value = box_end_row_data['values']
            else:
                range_obj.Value = None
            return range_obj

        return range_obj

    def cells_side_effect(row, col):
        cell = MagicMock()
        cell._row = row
        cell._col = col
        return cell

    worksheet.Range.side_effect = range_side_effect
    worksheet.Cells.side_effect = cells_side_effect

    return worksheet


class TestBatchReadBoxRanges(unittest.TestCase):

    def setUp(self):
        from excel_automation.box_list_export_config import BoxListExportConfig
        self.config = BoxListExportConfig()

    def test_returns_empty_for_unknown_size(self):
        from excel_automation.box_list_export_manager import BoxListExportManager

        size_col_data = ((38.0,), (40.0,), (42.0,))
        box_start_data = {'row': 15, 'values': ((1, 5, 10),)}
        box_end_data = {'row': 16, 'values': ((4, 9, 15),)}

        worksheet = make_mock_worksheet(
            size_col_data, box_start_data, box_end_data, {}
        )

        manager = BoxListExportManager(self.config)
        result = manager.read_box_ranges(worksheet, ["099"])
        self.assertEqual(result.get("099", []), [])

    def test_box_range_dataclass_is_valid(self):
        from excel_automation.box_list_export_manager import BoxRange
        br = BoxRange(sizes=["038"], box_start=1, box_end=5, column_number=7, total_pcs=10)
        self.assertTrue(br.is_valid())
        self.assertFalse(br.is_combined())
        self.assertEqual(br.get_box_numbers(), [1, 2, 3, 4, 5])

    def test_box_range_invalid_when_start_gt_end(self):
        from excel_automation.box_list_export_manager import BoxRange
        br = BoxRange(sizes=["038"], box_start=10, box_end=5, column_number=7)
        self.assertFalse(br.is_valid())

    def test_box_range_combined(self):
        from excel_automation.box_list_export_manager import BoxRange
        br = BoxRange(sizes=["038", "040"], box_start=1, box_end=3, column_number=7)
        self.assertTrue(br.is_combined())
        self.assertEqual(br.get_size_label("/"), "38/40")


class TestBatchWritePasteAndFormat(unittest.TestCase):

    def setUp(self):
        from excel_automation.box_list_export_config import BoxListExportConfig
        from excel_automation.box_list_export_manager import BoxListExportManager
        self.config = BoxListExportConfig()
        self.manager = BoxListExportManager(self.config)

    def _make_mock_sheet(self):
        sheet = MagicMock()
        sheet.Name = "TestSheet"

        cells_dict = {}

        def cells_side_effect(row, col):
            key = (row, col)
            if key not in cells_dict:
                cell = MagicMock()
                cell.Font = MagicMock()
                cells_dict[key] = cell
            return cells_dict[key]

        sheet.Cells.side_effect = cells_side_effect

        range_obj = MagicMock()
        range_obj.Font = MagicMock()
        sheet.Range.return_value = range_obj

        return sheet

    def _make_mock_workbook(self):
        wb = MagicMock()
        wb.Name = "Test.xlsx"
        return wb

    def _make_mock_source_worksheet(self):
        ws = MagicMock()
        ws.Cells.return_value.Value = "PO12345"
        return ws

    def test_paste_returns_true_on_success(self):
        from excel_automation.box_list_export_manager import BoxRange
        new_sheet = self._make_mock_sheet()
        workbook = self._make_mock_workbook()
        source_ws = self._make_mock_source_worksheet()

        box_ranges = [
            BoxRange(sizes=["038"], box_start=1, box_end=3, column_number=7, total_pcs=10),
        ]

        result = self.manager.paste_and_format_to_excel(
            workbook, source_ws, box_ranges, new_sheet, "A", 1, None
        )
        self.assertTrue(result)

    def test_paste_handles_empty_box_ranges(self):
        new_sheet = self._make_mock_sheet()
        workbook = self._make_mock_workbook()
        source_ws = self._make_mock_source_worksheet()

        result = self.manager.paste_and_format_to_excel(
            workbook, source_ws, [], new_sheet, "A", 1, None
        )
        self.assertTrue(result)

    def test_paste_uses_range_for_batch_write(self):
        from excel_automation.box_list_export_manager import BoxRange
        new_sheet = self._make_mock_sheet()
        workbook = self._make_mock_workbook()
        source_ws = self._make_mock_source_worksheet()

        box_ranges = [
            BoxRange(sizes=["038"], box_start=1, box_end=5, column_number=7, total_pcs=10),
        ]

        self.manager.paste_and_format_to_excel(
            workbook, source_ws, box_ranges, new_sheet, "A", 1, None
        )

        new_sheet.Range.assert_called()


class TestExportStepMethods(unittest.TestCase):

    def setUp(self):
        from excel_automation.box_list_export_config import BoxListExportConfig
        from excel_automation.box_list_export_manager import BoxListExportManager
        self.config = BoxListExportConfig()
        self.manager = BoxListExportManager(self.config)

    def test_generate_box_list_text_single_size(self):
        from excel_automation.box_list_export_manager import BoxRange
        box_ranges = [
            BoxRange(sizes=["038"], box_start=1, box_end=3, column_number=7),
        ]
        text = self.manager.generate_box_list_text(box_ranges)
        self.assertIn("SIZE 38", text)
        self.assertIn("1", text)
        self.assertIn("2", text)
        self.assertIn("3", text)

    def test_split_into_columns_respects_max_rows(self):
        lines = [f"line_{i}" for i in range(100)]
        columns = self.manager.split_into_columns(lines)
        max_content = self.config.get_max_rows_per_column() - self.config.get_header_rows()
        for col in columns:
            self.assertLessEqual(len(col), max_content)

    def test_build_export_result_success(self):
        from excel_automation.box_list_export_manager import BoxRange, BoxListExportResult
        box_ranges = [
            BoxRange(sizes=["038"], box_start=1, box_end=3, column_number=7),
        ]
        result = BoxListExportResult(
            success=True,
            text="test",
            box_ranges=box_ranges,
            total_boxes=3,
            header="test_header",
            total_columns=1
        )
        self.assertTrue(result.success)
        self.assertEqual(result.total_boxes, 3)
        summary = result.get_summary()
        self.assertIn("3 thùng", summary)


if __name__ == "__main__":
    unittest.main()
