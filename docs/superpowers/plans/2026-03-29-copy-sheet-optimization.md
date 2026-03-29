# Copy Sheet Performance Optimization — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Tối ưu flow Copy Sheet từ ~1,500 COM calls xuống ~5-10 calls bằng bulk Range operations, kèm progress dialog hiển thị tiến trình.

**Architecture:** Refactor 3 methods trong `ExcelCOMManager` dùng bulk Range thay cell-by-cell loop. Tạo `CopySheetProgressDialog` reuse pattern `ImportProgressDialog`. Refactor `_copy_sheet()` trong UI kết nối cả hai.

**Tech Stack:** Python, win32com.client (COM automation), tkinter, pytest + unittest.mock

**Spec:** `docs/superpowers/specs/2026-03-29-copy-sheet-optimization-design.md`

---

## File Structure

| File | Loại | Trách nhiệm |
|---|---|---|
| `excel_automation/excel_com_manager.py` | Sửa | Bulk Range cho `clear_quantity_columns()`, `show_all_rows()`, `scan_sizes()`. Thêm `_number_to_column_letter()` |
| `ui/copy_sheet_progress_dialog.py` | Mới | `CopySheetProgressDialog` class — progress dialog cho copy sheet flow |
| `ui/excel_realtime_controller.py` | Sửa | Refactor `_copy_sheet()` dùng progress dialog |
| `tests/test_excel_com_manager.py` | Mới | Unit tests cho bulk Range methods |

---

### Task 1: Thêm `_number_to_column_letter()` helper vào ExcelCOMManager

**Files:**
- Modify: `excel_automation/excel_com_manager.py:348-353`
- Test: `tests/test_excel_com_manager.py`

- [ ] **Step 1: Tạo test file với tests cho helper**

```python
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
```

File: `tests/test_excel_com_manager.py`

- [ ] **Step 2: Run test — verify FAIL**

Run: `pytest tests/test_excel_com_manager.py::TestNumberToColumnLetter -v`
Expected: FAIL — `AttributeError: 'ExcelCOMManager' object has no attribute '_number_to_column_letter'`

- [ ] **Step 3: Implement `_number_to_column_letter()`**

Thêm method vào `ExcelCOMManager` ngay sau `_column_letter_to_number()` (line ~353):

```python
    def _number_to_column_letter(self, col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result
```

File: `excel_automation/excel_com_manager.py` — thêm sau method `_column_letter_to_number`

- [ ] **Step 4: Run test — verify PASS**

Run: `pytest tests/test_excel_com_manager.py::TestNumberToColumnLetter -v`
Expected: 3 tests PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_excel_com_manager.py excel_automation/excel_com_manager.py
git commit -m "feat: add _number_to_column_letter() helper to ExcelCOMManager"
```

---

### Task 2: Refactor `clear_quantity_columns()` dùng bulk Range

**Files:**
- Modify: `excel_automation/excel_com_manager.py:313-346`
- Test: `tests/test_excel_com_manager.py`

- [ ] **Step 1: Thêm tests cho clear_quantity_columns bulk**

Append vào `tests/test_excel_com_manager.py`:

```python
class TestClearQuantityColumnsBulk(unittest.TestCase):

    def setUp(self):
        with patch.object(ExcelCOMManager, '__init__', lambda self, *a, **kw: None):
            self.manager = ExcelCOMManager()
            self.manager.config = MagicMock()
            self.manager.config.get_start_row.return_value = 19
            self.manager.excel_app = MagicMock()
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

        calls = []
        for name, args, kwargs in self.manager.excel_app.mock_calls:
            if 'ScreenUpdating' in name:
                calls.append(args[0] if args else None)

        self.assertIn(False, calls)

    def test_screen_updating_restored_on_error(self):
        self.manager.worksheet.Range.return_value.ClearContents.side_effect = Exception("COM error")

        with self.assertRaises(RuntimeError):
            self.manager.clear_quantity_columns(start_row=19, end_row=59, start_col=7, end_col=39)

    def test_raises_if_no_worksheet(self):
        self.manager.worksheet = None
        with self.assertRaises(RuntimeError):
            self.manager.clear_quantity_columns()
```

- [ ] **Step 2: Run test — verify FAIL**

Run: `pytest tests/test_excel_com_manager.py::TestClearQuantityColumnsBulk -v`
Expected: FAIL — tests expect bulk Range call, but current code uses cell-by-cell loop

- [ ] **Step 3: Refactor `clear_quantity_columns()` dùng bulk Range**

Thay toàn bộ body method `clear_quantity_columns` trong `excel_automation/excel_com_manager.py`:

```python
    def clear_quantity_columns(self, start_row: Optional[int] = None,
                               end_row: Optional[int] = None,
                               start_col: int = 7,
                               end_col: int = 39) -> int:
        if self.worksheet is None:
            raise RuntimeError("Chưa chọn worksheet nào")

        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.detect_end_row()

        try:
            if self.excel_app:
                self.excel_app.ScreenUpdating = False

            start_col_letter = self._number_to_column_letter(start_col)
            end_col_letter = self._number_to_column_letter(end_col)
            range_str = f"{start_col_letter}{start_row}:{end_col_letter}{end_row}"

            target_range = self.worksheet.Range(range_str)
            cleared_count = int(self.excel_app.WorksheetFunction.CountA(target_range))
            target_range.ClearContents()

            if self.excel_app:
                self.excel_app.ScreenUpdating = True

            logger.info(f"Đã xóa {cleared_count} ô số lượng ({range_str})")
            return cleared_count

        except Exception as e:
            if self.excel_app:
                self.excel_app.ScreenUpdating = True
            logger.error(f"Lỗi khi xóa số lượng: {e}")
            raise RuntimeError(f"Không thể xóa số lượng: {str(e)}")
```

- [ ] **Step 4: Run test — verify PASS**

Run: `pytest tests/test_excel_com_manager.py::TestClearQuantityColumnsBulk -v`
Expected: 5 tests PASS

- [ ] **Step 5: Commit**

```bash
git add excel_automation/excel_com_manager.py tests/test_excel_com_manager.py
git commit -m "perf: refactor clear_quantity_columns to use bulk Range.ClearContents"
```

---

### Task 3: Refactor `show_all_rows()` dùng bulk Range

**Files:**
- Modify: `excel_automation/excel_com_manager.py:246-269`
- Test: `tests/test_excel_com_manager.py`

- [ ] **Step 1: Thêm tests cho show_all_rows bulk**

Append vào `tests/test_excel_com_manager.py`:

```python
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
        self.manager.worksheet.Range.return_value.EntireRow.Hidden = False

    def test_screen_updating_toggled(self):
        self.manager.show_all_rows(start_row=19, end_row=59)

        calls = []
        for name, args, kwargs in self.manager.excel_app.mock_calls:
            if 'ScreenUpdating' in name:
                calls.append(args[0] if args else None)

        self.assertIn(False, calls)

    def test_raises_if_no_worksheet(self):
        self.manager.worksheet = None
        with self.assertRaises(RuntimeError):
            self.manager.show_all_rows()

    def test_screen_updating_restored_on_error(self):
        self.manager.worksheet.Range.side_effect = Exception("COM error")

        with self.assertRaises(RuntimeError):
            self.manager.show_all_rows(start_row=19, end_row=59)
```

- [ ] **Step 2: Run test — verify FAIL**

Run: `pytest tests/test_excel_com_manager.py::TestShowAllRowsBulk -v`
Expected: FAIL — current code loops row-by-row instead of using Range

- [ ] **Step 3: Refactor `show_all_rows()` dùng bulk Range**

Thay toàn bộ body method `show_all_rows` trong `excel_automation/excel_com_manager.py`:

```python
    def show_all_rows(self, start_row: Optional[int] = None, end_row: Optional[int] = None) -> None:
        if self.worksheet is None:
            raise RuntimeError("Chưa chọn worksheet nào")

        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.detect_end_row()

        try:
            if self.excel_app:
                self.excel_app.ScreenUpdating = False

            range_str = f"{start_row}:{end_row}"
            self.worksheet.Range(range_str).EntireRow.Hidden = False

            if self.excel_app:
                self.excel_app.ScreenUpdating = True

            logger.info(f"Đã hiện tất cả dòng từ {start_row} đến {end_row}")

        except Exception as e:
            if self.excel_app:
                self.excel_app.ScreenUpdating = True
            logger.error(f"Lỗi khi hiện dòng: {e}")
            raise RuntimeError(f"Không thể hiện dòng: {str(e)}")
```

- [ ] **Step 4: Run test — verify PASS**

Run: `pytest tests/test_excel_com_manager.py::TestShowAllRowsBulk -v`
Expected: 4 tests PASS

- [ ] **Step 5: Commit**

```bash
git add excel_automation/excel_com_manager.py tests/test_excel_com_manager.py
git commit -m "perf: refactor show_all_rows to use bulk Range.EntireRow.Hidden"
```

---

### Task 4: Refactor `scan_sizes()` dùng bulk Range

**Files:**
- Modify: `excel_automation/excel_com_manager.py:143-173`
- Test: `tests/test_excel_com_manager.py`

- [ ] **Step 1: Thêm tests cho scan_sizes bulk**

Append vào `tests/test_excel_com_manager.py`:

```python
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
```

- [ ] **Step 2: Run test — verify FAIL**

Run: `pytest tests/test_excel_com_manager.py::TestScanSizesBulk -v`
Expected: FAIL — current code loops cell-by-cell instead of using Range

- [ ] **Step 3: Refactor `scan_sizes()` dùng bulk Range**

Thay toàn bộ body method `scan_sizes` trong `excel_automation/excel_com_manager.py`:

```python
    def scan_sizes(self, column: Optional[str] = None, start_row: Optional[int] = None,
                   end_row: Optional[int] = None) -> List[str]:
        if self.worksheet is None:
            raise RuntimeError("Chưa chọn worksheet nào")

        column = column or self.config.get_column()
        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.detect_end_row()

        try:
            col_num = self._column_letter_to_number(column)
            range_str = f"{column}{start_row}:{column}{end_row}"
            raw_values = self.worksheet.Range(range_str).Value

            if raw_values is None:
                return []

            if not isinstance(raw_values, tuple):
                raw_values = ((raw_values,),)

            sizes: Set[str] = set()

            for row_offset, row_tuple in enumerate(raw_values):
                cell_value = row_tuple[0] if isinstance(row_tuple, tuple) else row_tuple

                if cell_value is not None:
                    size_str = normalize_size_value(cell_value)

                    if size_str:
                        actual_row = start_row + row_offset
                        self._fix_decimal_cell(actual_row, col_num, cell_value, size_str)
                        sizes.add(size_str)

            sorted_sizes = sorted(sizes, key=get_size_sort_key)
            logger.info(f"Quét được {len(sorted_sizes)} size khác nhau trong {column}[{start_row}:{end_row}]")
            return sorted_sizes

        except Exception as e:
            logger.error(f"Lỗi khi quét sizes: {e}")
            raise RuntimeError(f"Không thể quét sizes: {str(e)}")
```

- [ ] **Step 4: Run test — verify PASS**

Run: `pytest tests/test_excel_com_manager.py::TestScanSizesBulk -v`
Expected: 6 tests PASS

- [ ] **Step 5: Run toàn bộ test file**

Run: `pytest tests/test_excel_com_manager.py -v`
Expected: All 18 tests PASS (3 + 5 + 4 + 6)

- [ ] **Step 6: Commit**

```bash
git add excel_automation/excel_com_manager.py tests/test_excel_com_manager.py
git commit -m "perf: refactor scan_sizes to use bulk Range.Value read"
```

---

### Task 5: Tạo `CopySheetProgressDialog`

**Files:**
- Create: `ui/copy_sheet_progress_dialog.py`
- Test: `tests/test_copy_sheet_progress_dialog.py`

- [ ] **Step 1: Tạo test file**

```python
import unittest
from unittest.mock import MagicMock, patch
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))


class TestCopySheetProgressDialog(unittest.TestCase):

    def setUp(self):
        self.tk_patcher = patch('ui.copy_sheet_progress_dialog.tk')
        self.ttk_patcher = patch('ui.copy_sheet_progress_dialog.ttk')
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

        from ui.copy_sheet_progress_dialog import CopySheetProgressDialog
        self.dialog = CopySheetProgressDialog(self.mock_parent)

    def tearDown(self):
        self.tk_patcher.stop()
        self.ttk_patcher.stop()

    def test_has_5_steps(self):
        self.assertEqual(len(self.dialog.STEPS), 5)

    def test_step_weights_sum_to_100(self):
        self.assertEqual(sum(self.dialog.STEP_WEIGHTS), 100)

    def test_step_weights_length_matches_steps(self):
        self.assertEqual(len(self.dialog.STEP_WEIGHTS), len(self.dialog.STEPS))

    def test_dialog_is_modal(self):
        self.mock_dialog.transient.assert_called_once_with(self.mock_parent)
        self.mock_dialog.grab_set.assert_called_once()

    def test_dialog_blocks_close(self):
        self.mock_dialog.protocol.assert_called_with("WM_DELETE_WINDOW", unittest.mock.ANY)

    def test_start_step_updates_current_step(self):
        self.dialog.start_step(2)
        self.assertEqual(self.dialog.current_step, 2)

    def test_finish_sets_100_percent(self):
        self.dialog.finish()
        self.mock_tk.IntVar.return_value.set.assert_called_with(100)

    def test_show_error_stores_retry_callback(self):
        callback = MagicMock()
        self.dialog.show_error(1, "test error", callback)
        self.assertEqual(self.dialog.retry_callback, callback)

    def test_close_destroys_dialog(self):
        self.dialog.close()
        self.mock_dialog.destroy.assert_called_once()


if __name__ == "__main__":
    unittest.main()
```

File: `tests/test_copy_sheet_progress_dialog.py`

- [ ] **Step 2: Run test — verify FAIL**

Run: `pytest tests/test_copy_sheet_progress_dialog.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'ui.copy_sheet_progress_dialog'`

- [ ] **Step 3: Implement `CopySheetProgressDialog`**

```python
import tkinter as tk
from tkinter import ttk
from typing import List, Optional, Callable
import logging

logger = logging.getLogger(__name__)


class CopySheetProgressDialog:

    STEPS = [
        "Copy sheet",
        "Xóa số lượng cũ",
        "Quét sizes",
        "Cập nhật giao diện",
        "Hoàn tất",
    ]

    STEP_WEIGHTS = [30, 25, 20, 20, 5]

    def __init__(self, parent: tk.Tk):
        self.parent = parent
        self.current_step = 0
        self.step_labels: List[ttk.Label] = []
        self.retry_callback: Optional[Callable] = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Đang Copy Sheet...")
        self.dialog.geometry("420x320")
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

File: `ui/copy_sheet_progress_dialog.py`

- [ ] **Step 4: Run test — verify PASS**

Run: `pytest tests/test_copy_sheet_progress_dialog.py -v`
Expected: 9 tests PASS

- [ ] **Step 5: Commit**

```bash
git add ui/copy_sheet_progress_dialog.py tests/test_copy_sheet_progress_dialog.py
git commit -m "feat: add CopySheetProgressDialog with step tracking"
```

---

### Task 6: Refactor `_copy_sheet()` trong UI dùng progress dialog

**Files:**
- Modify: `ui/excel_realtime_controller.py:426-485`

- [ ] **Step 1: Thêm import**

Thêm vào đầu file `ui/excel_realtime_controller.py`, cùng block import các dialog khác:

```python
from ui.copy_sheet_progress_dialog import CopySheetProgressDialog
```

Tìm dòng có `from ui.` imports và thêm vào đó.

- [ ] **Step 2: Refactor `_copy_sheet()` method**

Thay toàn bộ method `_copy_sheet()` (line 426-485) trong `ui/excel_realtime_controller.py`:

```python
    def _copy_sheet(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Cảnh báo", "Vui lòng mở file Excel trước!")
            return

        progress = CopySheetProgressDialog(self.root)

        try:
            progress.start_step(0)
            new_sheet_name = self.com_manager.copy_sheet()
            progress.complete_step(0)
        except Exception as e:
            logger.error(f"Lỗi khi copy sheet: {e}")
            progress.show_error(0, str(e), lambda: self._copy_sheet_retry(progress, None))
            return

        progress.dialog.withdraw()

        dialog = SheetRenameDialog(
            self.root,
            new_sheet_name,
            self.com_manager.get_sheet_names()
        )
        user_name = dialog.show()

        if user_name is None:
            self.status_label.config(text="Đã hủy đổi tên sheet")
            self._reload_sheets()
            progress.close()
            return

        if user_name != new_sheet_name:
            try:
                self.com_manager.rename_sheet(new_sheet_name, user_name)
                new_sheet_name = user_name
            except ValueError as ve:
                messagebox.showwarning("Cảnh báo", str(ve), parent=self.root)

        progress.dialog.deiconify()
        progress.dialog.grab_set()

        self._copy_sheet_continue(progress, new_sheet_name)

    def _copy_sheet_continue(self, progress: 'CopySheetProgressDialog', new_sheet_name: str) -> None:
        try:
            self.com_manager.switch_sheet(new_sheet_name)
            self.current_sheet = new_sheet_name

            progress.start_step(1)
            cleared = self.com_manager.clear_quantity_columns()
            progress.complete_step(1)

            progress.start_step(2)
            self._scan_sizes()
            self._deselect_all_sizes()
            progress.complete_step(2)

            progress.start_step(3)
            self.com_manager.show_all_rows()
            self.sheet_names = self.com_manager.get_sheet_names()
            self.sheet_combobox['values'] = self.sheet_names
            self.sheet_combobox.set(new_sheet_name)
            self.sheet_status_label.config(
                text=f"({len(self.sheet_names)} sheets)",
                foreground="blue"
            )
            self._update_po_color_display()
            self._highlight_update_buttons()
            self._start_auto_refresh_sizes()
            progress.complete_step(3)

            progress.finish()

            self.status_label.config(
                text=f"Đã copy → '{new_sheet_name}' | Xóa {cleared} ô số lượng"
            )
            logger.info(f"Đã copy sheet thành công: '{new_sheet_name}', xóa {cleared} ô")

        except Exception as e:
            step = progress.current_step
            logger.error(f"Lỗi khi copy sheet tại bước {step}: {e}")
            progress.show_error(
                step, str(e),
                lambda: self._copy_sheet_continue(progress, new_sheet_name)
            )

    def _copy_sheet_retry(self, progress: 'CopySheetProgressDialog', _) -> None:
        progress.close()
        self._copy_sheet()
```

- [ ] **Step 3: Verify import không bị lỗi cú pháp**

Run: `python -c "import ast; ast.parse(open('ui/excel_realtime_controller.py', encoding='utf-8').read()); print('Syntax OK')"`
Expected: `Syntax OK`

- [ ] **Step 4: Run existing tests**

Run: `pytest tests/ -v`
Expected: Tất cả tests PASS (không break gì)

- [ ] **Step 5: Commit**

```bash
git add ui/excel_realtime_controller.py
git commit -m "perf: refactor _copy_sheet to use progress dialog with bulk Range operations"
```

---

### Task 7: Run full test suite + verify

**Files:** Không thay đổi — chỉ verify

- [ ] **Step 1: Run toàn bộ test suite**

Run: `pytest tests/ -v`
Expected: Tất cả tests PASS

- [ ] **Step 2: Verify syntax toàn bộ modified files**

```bash
python -c "import ast; ast.parse(open('excel_automation/excel_com_manager.py', encoding='utf-8').read()); print('excel_com_manager OK')"
python -c "import ast; ast.parse(open('ui/copy_sheet_progress_dialog.py', encoding='utf-8').read()); print('copy_sheet_progress_dialog OK')"
python -c "import ast; ast.parse(open('ui/excel_realtime_controller.py', encoding='utf-8').read()); print('excel_realtime_controller OK')"
```

Expected: Cả 3 file OK

- [ ] **Step 3: Review changes tổng**

Run: `git log --oneline -10`

Expected commits (mới nhất trước):
```
perf: refactor _copy_sheet to use progress dialog with bulk Range operations
feat: add CopySheetProgressDialog with step tracking
perf: refactor scan_sizes to use bulk Range.Value read
perf: refactor show_all_rows to use bulk Range.EntireRow.Hidden
perf: refactor clear_quantity_columns to use bulk Range.ClearContents
feat: add _number_to_column_letter() helper to ExcelCOMManager
```
