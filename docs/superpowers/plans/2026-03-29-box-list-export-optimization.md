# Box List Export Optimization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Tối ưu hiệu năng xuất danh sách thùng bằng COM batch read/write và thêm progress dialog để UI không bị đóng băng.

**Architecture:** Thay thế cell-by-cell COM calls bằng Range-based batch operations trong `BoxListExportManager`. Tạo `BoxListExportProgressDialog` theo pattern `CopySheetProgressDialog`. Refactor `_export_box_list()` trong UI để dùng progress dialog với step-based flow và retry.

**Tech Stack:** Python, win32com.client (COM automation), tkinter, unittest/pytest

**Spec:** `docs/superpowers/specs/2026-03-29-box-list-export-optimization-design.md`

---

## File Structure

| File | Action | Responsibility |
|------|--------|---------------|
| `excel_automation/box_list_export_manager.py` | Modify | Tối ưu `read_box_ranges()` batch read, `paste_and_format_to_excel()` batch write |
| `ui/box_list_export_progress_dialog.py` | Create | Progress dialog với 6 steps |
| `ui/excel_realtime_controller.py` | Modify | Refactor `_export_box_list()` dùng progress dialog |
| `tests/test_box_list_export_progress_dialog.py` | Create | Unit tests cho progress dialog |
| `tests/test_box_list_batch_operations.py` | Create | Unit tests cho batch read/write |

---

### Task 1: Tạo BoxListExportProgressDialog

**Files:**
- Create: `ui/box_list_export_progress_dialog.py`
- Create: `tests/test_box_list_export_progress_dialog.py`
- Reference: `ui/copy_sheet_progress_dialog.py` (pattern mẫu)

- [ ] **Step 1: Viết test cho progress dialog**

```python
# tests/test_box_list_export_progress_dialog.py
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
```

- [ ] **Step 2: Chạy test xác nhận fail**

Run: `pytest tests/test_box_list_export_progress_dialog.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'ui.box_list_export_progress_dialog'`

- [ ] **Step 3: Implement BoxListExportProgressDialog**

```python
# ui/box_list_export_progress_dialog.py
import tkinter as tk
from tkinter import ttk
from typing import List, Optional, Callable
import logging

logger = logging.getLogger(__name__)


class BoxListExportProgressDialog:

    STEPS = [
        "Đọc dữ liệu thùng từ Excel",
        "Phân tích & gộp sizes",
        "Tạo sheet mới",
        "Ghi danh sách thùng vào sheet",
        "Copy vào clipboard",
        "Hoàn tất",
    ]

    STEP_WEIGHTS = [30, 15, 15, 25, 10, 5]

    def __init__(self, parent: tk.Tk):
        self.parent = parent
        self.current_step = 0
        self.step_labels: List[ttk.Label] = []
        self.retry_callback: Optional[Callable] = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Đang xuất danh sách thùng...")
        self.dialog.geometry("420x350")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)

        self._create_widgets()
        self._center_window()

    def _center_window(self) -> None:
        self.dialog.update_idletasks()
        w = self.dialog.winfo_width()
        h = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (w // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (h // 2)
        self.dialog.geometry(f"+{x}+{y}")

    def _create_widgets(self) -> None:
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(
            main_frame, variable=self.progress_var,
            maximum=100, length=350, mode='determinate'
        )
        self.progress_bar.pack(pady=(0, 5))

        self.percent_label = ttk.Label(main_frame, text="0%", font=("", 11, "bold"))
        self.percent_label.pack(pady=(0, 15))

        steps_frame = ttk.Frame(main_frame)
        steps_frame.pack(fill=tk.BOTH, expand=True)

        for step_text in self.STEPS:
            label = ttk.Label(steps_frame, text=f"  ⬚  {step_text}", foreground="gray")
            label.pack(anchor=tk.W, pady=2)
            self.step_labels.append(label)

        self.btn_frame = ttk.Frame(main_frame)
        self.btn_frame.pack(fill=tk.X, pady=(15, 0))

        self.error_label = ttk.Label(main_frame, text="", foreground="#c62828", wraplength=360)

    def start_step(self, step_index: int) -> None:
        self.current_step = step_index
        percent = sum(self.STEP_WEIGHTS[:step_index])
        self.progress_var.set(percent)
        self.percent_label.configure(text=f"{percent}%")

        for i, label in enumerate(self.step_labels):
            if i < step_index:
                label.configure(text=f"  ✅  {self.STEPS[i]}", foreground="#2e7d32")
            elif i == step_index:
                label.configure(text=f"  🔄  Đang {self.STEPS[i].lower()}...", foreground="#1565c0")
            else:
                label.configure(text=f"  ⬚  {self.STEPS[i]}", foreground="gray")

        self.dialog.update()

    def complete_step(self, step_index: int) -> None:
        self.step_labels[step_index].configure(
            text=f"  ✅  {self.STEPS[step_index]}", foreground="#2e7d32"
        )
        percent = sum(self.STEP_WEIGHTS[:step_index + 1])
        self.progress_var.set(percent)
        self.percent_label.configure(text=f"{percent}%")
        self.dialog.update()

    def finish(self) -> None:
        self.progress_var.set(100)
        self.percent_label.configure(text="100%")
        for i, label in enumerate(self.step_labels):
            label.configure(text=f"  ✅  {self.STEPS[i]}", foreground="#2e7d32")
        self.dialog.update()
        self.parent.after(1000, self.dialog.destroy)

    def show_error(self, step_index: int, error_msg: str, retry_callback: Callable) -> None:
        self.step_labels[step_index].configure(
            text=f"  ❌  {self.STEPS[step_index]}", foreground="#c62828"
        )
        self.error_label.configure(text=f"Lỗi: {error_msg}")
        self.error_label.pack(pady=(10, 0))

        self.retry_callback = retry_callback

        for widget in self.btn_frame.winfo_children():
            widget.destroy()

        ttk.Button(
            self.btn_frame, text="🔄 Thử lại",
            command=self._retry, width=15
        ).pack(side=tk.LEFT)
        ttk.Button(
            self.btn_frame, text="Đóng",
            command=self.dialog.destroy, width=15
        ).pack(side=tk.RIGHT)

        self.dialog.protocol("WM_DELETE_WINDOW", self.dialog.destroy)
        self.dialog.update()

    def close(self) -> None:
        self.dialog.destroy()

    def _retry(self) -> None:
        self.error_label.pack_forget()
        for widget in self.btn_frame.winfo_children():
            widget.destroy()
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        if self.retry_callback:
            self.retry_callback()
```

- [ ] **Step 4: Chạy test xác nhận pass**

Run: `pytest tests/test_box_list_export_progress_dialog.py -v`
Expected: All 15 tests PASS

- [ ] **Step 5: Commit**

```bash
git add ui/box_list_export_progress_dialog.py tests/test_box_list_export_progress_dialog.py
git commit -m "feat: add BoxListExportProgressDialog with 6-step progress tracking"
```

---

### Task 2: Tối ưu batch read trong `read_box_ranges()`

**Files:**
- Modify: `excel_automation/box_list_export_manager.py:94-184` (method `read_box_ranges`)
- Create: `tests/test_box_list_batch_operations.py`

- [ ] **Step 1: Viết test cho batch read**

```python
# tests/test_box_list_batch_operations.py
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


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy test xác nhận pass (test hiện tại chạy trên dataclass không đổi)**

Run: `pytest tests/test_box_list_batch_operations.py -v`
Expected: 4 tests PASS (dataclass tests chạy được, mock test cho unknown size cũng pass vì logic trả empty list cho size không tìm thấy)

- [ ] **Step 3: Implement batch read — thay thế `read_box_ranges()`**

Mở file `excel_automation/box_list_export_manager.py`, thay thế toàn bộ method `read_box_ranges` (dòng 94-184) bằng:

```python
    def read_box_ranges(
        self,
        worksheet: CDispatch,
        selected_sizes: List[str]
    ) -> Dict[str, List[Tuple[int, int, int, int]]]:
        box_start_row = self.config.get_box_start_row()
        box_end_row = self.config.get_box_end_row()
        size_column = self.config.get_size_column()
        size_data_start_row = self.config.get_size_data_start_row()

        size_column_number = self._column_letter_to_number(size_column)

        try:
            used_range = worksheet.UsedRange
            max_column = used_range.Column + used_range.Columns.Count - 1
            scan_end_column = max(39, max_column + 1)
        except Exception:
            scan_end_column = 39

        size_data_end_row = self.config.get_size_data_end_row()
        try:
            range_str = f"{size_column}{size_data_start_row}:{size_column}{size_data_end_row}"
            raw_size_values = worksheet.Range(range_str).Value
        except Exception:
            raw_size_values = None

        size_to_row: Dict[str, int] = {}
        if raw_size_values is not None:
            if not isinstance(raw_size_values, tuple):
                raw_size_values = ((raw_size_values,),)
            for row_offset, row_tuple in enumerate(raw_size_values):
                cell_value = row_tuple[0] if isinstance(row_tuple, tuple) else row_tuple
                if cell_value is not None and str(cell_value).strip() != "":
                    size_str = normalize_size_value(cell_value)
                    if size_str:
                        size_to_row[size_str] = size_data_start_row + row_offset

        start_col = 7
        end_col = scan_end_column - 1

        try:
            box_start_values_raw = worksheet.Range(
                worksheet.Cells(box_start_row, start_col),
                worksheet.Cells(box_start_row, end_col)
            ).Value
        except Exception:
            box_start_values_raw = None

        try:
            box_end_values_raw = worksheet.Range(
                worksheet.Cells(box_end_row, start_col),
                worksheet.Cells(box_end_row, end_col)
            ).Value
        except Exception:
            box_end_values_raw = None

        if box_start_values_raw is not None and not isinstance(box_start_values_raw, tuple):
            box_start_values_raw = ((box_start_values_raw,),)
        if box_end_values_raw is not None and not isinstance(box_end_values_raw, tuple):
            box_end_values_raw = ((box_end_values_raw,),)

        box_start_values = box_start_values_raw[0] if box_start_values_raw else ()
        box_end_values = box_end_values_raw[0] if box_end_values_raw else ()

        box_ranges: Dict[str, List[Tuple[int, int, int, int]]] = {}

        for size in selected_sizes:
            if size not in size_to_row:
                logger.warning(f"Size {size} không tìm thấy trong cột {size_column}")
                box_ranges[size] = []
                continue

            size_row = size_to_row[size]

            try:
                size_row_values_raw = worksheet.Range(
                    worksheet.Cells(size_row, start_col),
                    worksheet.Cells(size_row, end_col)
                ).Value
            except Exception:
                box_ranges[size] = []
                continue

            if size_row_values_raw is not None and not isinstance(size_row_values_raw, tuple):
                size_row_values_raw = ((size_row_values_raw,),)

            size_row_values = size_row_values_raw[0] if size_row_values_raw else ()

            size_box_ranges: List[Tuple[int, int, int, int]] = []

            for col_offset in range(len(size_row_values)):
                column_number = start_col + col_offset
                quantity_value = size_row_values[col_offset]

                if quantity_value is None:
                    continue

                try:
                    quantity = int(float(quantity_value))
                    if quantity <= 0:
                        continue
                except (ValueError, TypeError):
                    continue

                if col_offset >= len(box_start_values) or col_offset >= len(box_end_values):
                    continue

                box_start_value = box_start_values[col_offset]
                box_end_value = box_end_values[col_offset]

                if box_start_value is None or box_end_value is None:
                    continue

                try:
                    box_start = int(box_start_value)
                    box_end = int(box_end_value)
                except (ValueError, TypeError):
                    logger.warning(
                        f"Size {size}, cột {column_number}: box_start hoặc box_end không hợp lệ"
                    )
                    continue

                if box_start > box_end:
                    logger.warning(
                        f"Size {size}, cột {column_number}: box_start ({box_start}) > box_end ({box_end})"
                    )
                    continue

                size_box_ranges.append((box_start, box_end, column_number, quantity))

            box_ranges[size] = size_box_ranges

        return box_ranges
```

Lưu ý: Xóa import `find_last_data_row` khỏi dòng import ở đầu file vì không còn dùng:

Thay dòng 8:
```python
from excel_automation.utils import get_size_sort_key, normalize_size_value, find_last_data_row
```
thành:
```python
from excel_automation.utils import get_size_sort_key, normalize_size_value
```

- [ ] **Step 4: Chạy test xác nhận pass**

Run: `pytest tests/test_box_list_batch_operations.py -v`
Expected: All 4 tests PASS

- [ ] **Step 5: Chạy toàn bộ test suite kiểm tra regression**

Run: `pytest tests/ -v`
Expected: Tất cả tests PASS, không có regression

- [ ] **Step 6: Commit**

```bash
git add excel_automation/box_list_export_manager.py tests/test_box_list_batch_operations.py
git commit -m "perf: batch read COM operations in read_box_ranges - reduce ~960 COM calls to ~3+N"
```

---

### Task 3: Tối ưu batch write trong `paste_and_format_to_excel()`

**Files:**
- Modify: `excel_automation/box_list_export_manager.py:347-400` (method `paste_and_format_to_excel`)

- [ ] **Step 1: Viết test cho batch write**

Thêm vào file `tests/test_box_list_batch_operations.py`:

```python
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

    def test_paste_writes_header_to_first_cell(self):
        from excel_automation.box_list_export_manager import BoxRange
        new_sheet = self._make_mock_sheet()
        workbook = self._make_mock_workbook()
        source_ws = self._make_mock_source_worksheet()

        box_ranges = [
            BoxRange(sizes=["038"], box_start=1, box_end=2, column_number=7, total_pcs=10),
        ]

        self.manager.paste_and_format_to_excel(
            workbook, source_ws, box_ranges, new_sheet, "A", 1, None
        )

        new_sheet.Cells(1, 1).Font.Bold = True

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
```

- [ ] **Step 2: Chạy test xác nhận pass (test hiện tại pass vì method chưa đổi)**

Run: `pytest tests/test_box_list_batch_operations.py -v`
Expected: All 8 tests PASS

- [ ] **Step 3: Implement batch write — thay thế `paste_and_format_to_excel()`**

Mở file `excel_automation/box_list_export_manager.py`, thay thế toàn bộ method `paste_and_format_to_excel` (dòng 347-400) bằng:

```python
    def paste_and_format_to_excel(
        self,
        workbook: CDispatch,
        worksheet: CDispatch,
        box_ranges: List[BoxRange],
        new_sheet: CDispatch,
        start_column: str = "A",
        start_row: int = 1,
        items_per_box: Optional[int] = None
    ) -> bool:
        try:
            header = self.generate_header(workbook, worksheet, items_per_box)
            header_rows = self.config.get_header_rows()

            content_lines = []
            for box_range in box_ranges:
                separator = self.config.get_combined_size_separator()
                size_label = box_range.get_size_label(separator)
                content_lines.append(f"SIZE {size_label}")

                for box_number in box_range.get_box_numbers():
                    content_lines.append(str(box_number))

            columns = self.split_into_columns(content_lines)
            start_col_num = self._column_letter_to_number(start_column)

            new_sheet.Cells(start_row, start_col_num).Value = header
            new_sheet.Cells(start_row, start_col_num).Font.Bold = True
            new_sheet.Cells(start_row, start_col_num).Font.Size = 20
            new_sheet.Cells(start_row, start_col_num).HorizontalAlignment = -4131

            for col_idx, column_lines in enumerate(columns):
                col_num = start_col_num + col_idx
                data_start_row = start_row + header_rows

                if column_lines:
                    data_array = [[line] for line in column_lines]
                    data_end_row = data_start_row + len(column_lines) - 1
                    new_sheet.Range(
                        new_sheet.Cells(data_start_row, col_num),
                        new_sheet.Cells(data_end_row, col_num)
                    ).Value = data_array

                    new_sheet.Range(
                        new_sheet.Cells(data_start_row, col_num),
                        new_sheet.Cells(data_end_row, col_num)
                    ).HorizontalAlignment = -4108

                    bold_rows = []
                    non_bold_rows = []
                    for line_idx, line in enumerate(column_lines):
                        row_num = data_start_row + line_idx
                        if line.startswith("SIZE "):
                            bold_rows.append(row_num)
                        else:
                            non_bold_rows.append(row_num)

                    if bold_rows:
                        bold_range = self._build_union_range(
                            new_sheet, bold_rows, col_num
                        )
                        if bold_range:
                            bold_range.Font.Bold = True

                    if non_bold_rows:
                        non_bold_range = self._build_union_range(
                            new_sheet, non_bold_rows, col_num
                        )
                        if non_bold_range:
                            non_bold_range.Font.Bold = False

            logger.info(f"Đã paste và format {len(columns)} cột vào sheet mới: {new_sheet.Name}")
            return True
        except Exception as e:
            logger.error(f"Lỗi khi paste vào Excel: {e}", exc_info=True)
            return False
```

Thêm method helper `_build_union_range` vào cuối class `BoxListExportManager` (trước `_column_letter_to_number`):

```python
    def _build_union_range(
        self,
        sheet: CDispatch,
        rows: List[int],
        col_num: int
    ) -> Optional[CDispatch]:
        if not rows:
            return None

        try:
            excel_app = sheet.Application
            result_range = sheet.Cells(rows[0], col_num)

            batch_size = 30
            for i in range(1, len(rows), batch_size):
                batch = rows[i:i + batch_size]
                for row in batch:
                    result_range = excel_app.Union(
                        result_range, sheet.Cells(row, col_num)
                    )

            return result_range
        except Exception as e:
            logger.warning(f"Không thể tạo union range, fallback từng cell: {e}")
            return None
```

- [ ] **Step 4: Chạy test xác nhận pass**

Run: `pytest tests/test_box_list_batch_operations.py -v`
Expected: All 8 tests PASS

- [ ] **Step 5: Chạy toàn bộ test suite kiểm tra regression**

Run: `pytest tests/ -v`
Expected: Tất cả tests PASS

- [ ] **Step 6: Commit**

```bash
git add excel_automation/box_list_export_manager.py tests/test_box_list_batch_operations.py
git commit -m "perf: batch write COM operations in paste_and_format_to_excel - reduce ~600 COM calls to ~5-10"
```

---

### Task 4: Tách `export_box_list()` thành các step riêng biệt

**Files:**
- Modify: `excel_automation/box_list_export_manager.py:402-485` (method `export_box_list`)

Hiện tại `export_box_list()` gộp tất cả logic (read + detect + generate text + clipboard) vào 1 method. Cần tách ra để UI có thể gọi từng step với progress dialog.

- [ ] **Step 1: Viết test cho step methods**

Thêm vào file `tests/test_box_list_batch_operations.py`:

```python
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
```

- [ ] **Step 2: Chạy test xác nhận pass**

Run: `pytest tests/test_box_list_batch_operations.py::TestExportStepMethods -v`
Expected: All 3 tests PASS

- [ ] **Step 3: Refactor `export_box_list()` — tách thành step methods**

Mở file `excel_automation/box_list_export_manager.py`, thay thế method `export_box_list` bằng:

```python
    def step_read_box_ranges(
        self,
        worksheet: CDispatch,
        selected_sizes: List[str]
    ) -> Dict[str, List[Tuple[int, int, int, int]]]:
        return self.read_box_ranges(worksheet, selected_sizes)

    def step_analyze_and_build_result(
        self,
        workbook: CDispatch,
        worksheet: CDispatch,
        selected_sizes: List[str],
        box_ranges_dict: Dict[str, List[Tuple[int, int, int, int]]],
        items_per_box: Optional[int] = None
    ) -> BoxListExportResult:
        valid_count = sum(
            1 for size_ranges in box_ranges_dict.values()
            if len(size_ranges) > 0
        )

        if valid_count == 0:
            return BoxListExportResult(
                success=False,
                error_message="Không có size nào có dữ liệu box hợp lệ"
            )

        box_ranges = self.detect_combined_sizes(
            selected_sizes, box_ranges_dict, items_per_box
        )

        partial_count = sum(1 for br in box_ranges if br.is_partial())
        logger.info(
            f"Đã phát hiện {len(box_ranges)} box ranges "
            f"({sum(1 for br in box_ranges if br.is_combined())} kết hợp, "
            f"{partial_count} thùng lẻ)"
        )

        header = self.generate_header(workbook, worksheet, items_per_box)

        content_lines = []
        for box_range in box_ranges:
            separator = self.config.get_combined_size_separator()
            size_label = box_range.get_size_label(separator)
            content_lines.append(f"SIZE {size_label}")
            for box_number in box_range.get_box_numbers():
                content_lines.append(str(box_number))

        columns = self.split_into_columns(content_lines)
        total_columns = len(columns)

        text_parts = []
        for col_idx, column_lines in enumerate(columns):
            text_parts.append(header)
            text_parts.append("")
            text_parts.extend(column_lines)
            if col_idx < len(columns) - 1:
                text_parts.append("")

        text = "\n".join(text_parts)
        total_boxes = sum(len(br.get_box_numbers()) for br in box_ranges)

        return BoxListExportResult(
            success=True,
            text=text,
            box_ranges=box_ranges,
            total_boxes=total_boxes,
            header=header,
            total_columns=total_columns
        )

    def export_box_list(
        self,
        excel_app: CDispatch,
        workbook: CDispatch,
        worksheet: CDispatch,
        selected_sizes: List[str],
        items_per_box: Optional[int] = None
    ) -> BoxListExportResult:
        logger.info(f"Bắt đầu xuất danh sách thùng cho {len(selected_sizes)} sizes")

        try:
            excel_app.ScreenUpdating = False

            box_ranges_dict = self.step_read_box_ranges(worksheet, selected_sizes)

            result = self.step_analyze_and_build_result(
                workbook, worksheet, selected_sizes,
                box_ranges_dict, items_per_box
            )

            if result.success:
                self.copy_to_clipboard(result.text)
                logger.info(
                    f"Hoàn thành xuất danh sách thùng: "
                    f"{result.total_boxes} thùng, {result.total_columns} cột"
                )

            return result

        except Exception as e:
            error_msg = f"Lỗi khi xuất danh sách thùng: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return BoxListExportResult(success=False, error_message=error_msg)
        finally:
            excel_app.ScreenUpdating = True
```

- [ ] **Step 4: Chạy toàn bộ test suite kiểm tra regression**

Run: `pytest tests/ -v`
Expected: Tất cả tests PASS

- [ ] **Step 5: Commit**

```bash
git add excel_automation/box_list_export_manager.py tests/test_box_list_batch_operations.py
git commit -m "refactor: split export_box_list into step methods for progress dialog integration"
```

---

### Task 5: Refactor `_export_box_list()` trong UI dùng progress dialog

**Files:**
- Modify: `ui/excel_realtime_controller.py:1538-1631` (method `_export_box_list`)
- Import: `ui/box_list_export_progress_dialog.py`

- [ ] **Step 1: Thêm import BoxListExportProgressDialog**

Mở file `ui/excel_realtime_controller.py`, tìm dòng import hiện tại:

```python
from excel_automation.box_list_export_config import BoxListExportConfig
from excel_automation.box_list_export_manager import BoxListExportManager
```

Thêm ngay sau:

```python
from ui.box_list_export_progress_dialog import BoxListExportProgressDialog
```

- [ ] **Step 2: Thay thế method `_export_box_list()`**

Mở file `ui/excel_realtime_controller.py`, thay thế toàn bộ method `_export_box_list` (dòng 1538-1631) bằng:

```python
    def _export_box_list(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Cảnh báo", "Vui lòng mở file Excel trước!")
            return

        selected_sizes = [
            size for size, var in self.checkboxes.items()
            if var.get()
        ]

        if not selected_sizes:
            messagebox.showwarning(
                "Cảnh báo",
                "Vui lòng chọn ít nhất một size để xuất danh sách thùng!"
            )
            return

        config = BoxListExportConfig()
        manager = BoxListExportManager(config)
        items_per_box = self._extract_items_per_box()

        progress = BoxListExportProgressDialog(self.root)

        box_ranges_dict = None
        result = None
        new_sheet = None

        def run_export_steps(start_from: int = 0):
            nonlocal box_ranges_dict, result, new_sheet

            try:
                if start_from <= 0:
                    progress.start_step(0)
                    self.com_manager.excel_app.ScreenUpdating = False
                    box_ranges_dict = manager.step_read_box_ranges(
                        self.com_manager.worksheet, selected_sizes
                    )
                    progress.complete_step(0)

                if start_from <= 1:
                    progress.start_step(1)
                    result = manager.step_analyze_and_build_result(
                        self.com_manager.workbook,
                        self.com_manager.worksheet,
                        selected_sizes,
                        box_ranges_dict,
                        items_per_box
                    )

                    if not result.success:
                        self.com_manager.excel_app.ScreenUpdating = True
                        progress.close()
                        messagebox.showerror(
                            "Lỗi",
                            f"Không thể xuất danh sách thùng:\n\n{result.error_message}"
                        )
                        self.status_label.config(text="Lỗi khi xuất danh sách thùng")
                        return

                    progress.complete_step(1)

                if start_from <= 2:
                    progress.start_step(2)
                    new_sheet = manager.create_new_sheet(
                        self.com_manager.workbook,
                        self.com_manager.worksheet
                    )
                    progress.complete_step(2)

                if start_from <= 3:
                    progress.start_step(3)
                    manager.paste_and_format_to_excel(
                        self.com_manager.workbook,
                        self.com_manager.worksheet,
                        result.box_ranges,
                        new_sheet,
                        "A",
                        1,
                        items_per_box
                    )
                    progress.complete_step(3)

                if start_from <= 4:
                    progress.start_step(4)
                    manager.copy_to_clipboard(result.text)
                    progress.complete_step(4)

                self.com_manager.excel_app.ScreenUpdating = True

                progress.start_step(5)
                progress.complete_step(5)
                progress.finish()

                summary = result.get_summary()
                self.status_label.config(text=summary)
                logger.info(f"Xuất danh sách thùng thành công: {summary}")

                self.root.after(1200, lambda: messagebox.showinfo(
                    "Thành Công",
                    f"{summary}\n\n"
                    f"Danh sách thùng đã được xuất vào sheet mới: {new_sheet.Name}\n"
                    f"Tất cả nội dung đã được căn giữa tự động."
                ))

            except Exception as e:
                self.com_manager.excel_app.ScreenUpdating = True
                step = progress.current_step
                logger.error(f"Lỗi khi xuất danh sách thùng tại bước {step}: {e}", exc_info=True)
                progress.show_error(step, str(e), lambda: run_export_steps(step))

        run_export_steps()
```

- [ ] **Step 3: Chạy toàn bộ test suite kiểm tra regression**

Run: `pytest tests/ -v`
Expected: Tất cả tests PASS

- [ ] **Step 4: Commit**

```bash
git add ui/excel_realtime_controller.py
git commit -m "feat: integrate BoxListExportProgressDialog into box list export flow"
```

---

### Task 6: Test thủ công end-to-end

**Files:** Không thay đổi — chỉ verify

- [ ] **Step 1: Chạy toàn bộ test suite lần cuối**

Run: `pytest tests/ -v`
Expected: Tất cả tests PASS, không có regression

- [ ] **Step 2: Kiểm tra app chạy không lỗi import**

Run: `python -c "from ui.box_list_export_progress_dialog import BoxListExportProgressDialog; print('OK')"`
Expected: `OK`

Run: `python -c "from excel_automation.box_list_export_manager import BoxListExportManager; print('OK')"`
Expected: `OK`

- [ ] **Step 3: Ghi chú test thủ công cần thực hiện**

Test thủ công cần thực hiện trên máy Windows có Excel:
1. Mở app, load file Excel packing list
2. Chọn vài sizes, bấm xuất danh sách thùng
3. Xác nhận: progress dialog hiện đúng, các step chạy lần lượt
4. Xác nhận: sheet mới tạo đúng format (header bold size 20 căn trái, data căn giữa, SIZE lines bold)
5. Xác nhận: clipboard có nội dung đúng
6. Test retry: ngắt kết nối Excel giữa chừng → dialog hiện lỗi + nút retry
7. So sánh tốc độ trước/sau: với 10+ sizes, thời gian phải giảm rõ rệt

- [ ] **Step 4: Commit final**

```bash
git add -A
git commit -m "docs: complete box list export optimization implementation"
```
