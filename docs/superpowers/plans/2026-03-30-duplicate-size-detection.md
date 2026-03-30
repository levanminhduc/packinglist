# Duplicate Size Detection Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Tự động phát hiện size trùng sau khi copy sheet, hiện dialog cho user chọn giữ/xóa dòng.

**Architecture:** Tách business logic (`DuplicateSizeDetector`) khỏi UI (`DuplicateSizeDialog`). Detector dùng COM Range read + `normalize_size_value` để build size→rows mapping, lọc ra size có >= 2 dòng. Dialog hiện grouped checkboxes, user chọn giữ, dòng không check bị xóa bottom-up. Tích hợp vào `_copy_sheet_continue()` sau step 2.

**Tech Stack:** Python, tkinter, win32com.client (COM automation), unittest + MagicMock

**Spec:** `docs/superpowers/specs/2026-03-30-duplicate-size-detection-design.md`

---

## File Structure

| Action | File | Responsibility |
|--------|------|----------------|
| Create | `excel_automation/duplicate_size_detector.py` | Business logic: detect duplicate sizes, delete rows via COM |
| Create | `ui/duplicate_size_dialog.py` | tkinter dialog: hiện grouped checkboxes cho user chọn giữ/xóa |
| Create | `tests/test_duplicate_size_detector.py` | Unit tests cho DuplicateSizeDetector |
| Modify | `excel_automation/__init__.py` | Export `DuplicateSizeDetector` |
| Modify | `ui/excel_realtime_controller.py` | Gọi detect + dialog sau copy sheet |

---

## Task 1: DuplicateSizeDetector — detect_duplicates

**Files:**
- Create: `tests/test_duplicate_size_detector.py`
- Create: `excel_automation/duplicate_size_detector.py`

- [ ] **Step 1: Write failing test for detect_duplicates — no duplicates**

File: `tests/test_duplicate_size_detector.py`

```python
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


if __name__ == '__main__':
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_duplicate_size_detector.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'excel_automation.duplicate_size_detector'`

- [ ] **Step 3: Write minimal implementation of DuplicateSizeDetector.detect_duplicates**

File: `excel_automation/duplicate_size_detector.py`

```python
from typing import Dict, List, Optional
import logging

from excel_automation.utils import normalize_size_value, get_size_sort_key

logger = logging.getLogger(__name__)


class DuplicateSizeDetector:

    def __init__(self, com_manager):
        self.com_manager = com_manager

    def detect_duplicates(
        self,
        column: Optional[str] = None,
        start_row: Optional[int] = None,
        end_row: Optional[int] = None
    ) -> Dict[str, List[int]]:
        worksheet = self.com_manager.worksheet
        if worksheet is None:
            return {}

        column = column or self.com_manager.config.get_column()
        start_row = start_row or self.com_manager.config.get_start_row()
        end_row = end_row or self.com_manager.detect_end_row()

        try:
            range_str = f"{column}{start_row}:{column}{end_row}"
            raw_values = worksheet.Range(range_str).Value

            if raw_values is None:
                return {}

            if not isinstance(raw_values, tuple):
                raw_values = ((raw_values,),)

            size_rows: Dict[str, List[int]] = {}

            for row_offset, row_tuple in enumerate(raw_values):
                cell_value = row_tuple[0] if isinstance(row_tuple, tuple) else row_tuple

                if cell_value is not None:
                    size_str = normalize_size_value(cell_value)
                    if size_str:
                        actual_row = start_row + row_offset
                        if size_str not in size_rows:
                            size_rows[size_str] = []
                        size_rows[size_str].append(actual_row)

            duplicates = {
                size: rows
                for size, rows in size_rows.items()
                if len(rows) >= 2
            }

            if duplicates:
                logger.info(
                    f"Phát hiện {len(duplicates)} size trùng: "
                    f"{', '.join(f'{s}({len(r)} dòng)' for s, r in duplicates.items())}"
                )
            else:
                logger.info("Không phát hiện size trùng")

            return duplicates

        except Exception as e:
            logger.error(f"Lỗi khi detect size trùng: {e}")
            raise RuntimeError(f"Không thể quét size trùng: {str(e)}")
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_duplicate_size_detector.py -v`
Expected: 4 tests PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_duplicate_size_detector.py excel_automation/duplicate_size_detector.py
git commit -m "feat: add DuplicateSizeDetector.detect_duplicates with tests"
```

---

## Task 2: DuplicateSizeDetector — delete_rows

**Files:**
- Modify: `tests/test_duplicate_size_detector.py`
- Modify: `excel_automation/duplicate_size_detector.py`

- [ ] **Step 1: Write failing tests for delete_rows**

Thêm vào cuối file `tests/test_duplicate_size_detector.py`, trước `if __name__`:

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_duplicate_size_detector.py::TestDeleteRows -v`
Expected: FAIL with `AttributeError: 'DuplicateSizeDetector' object has no attribute 'delete_rows'`

- [ ] **Step 3: Add delete_rows method**

Thêm vào cuối class `DuplicateSizeDetector` trong `excel_automation/duplicate_size_detector.py`:

```python
    def delete_rows(self, rows_to_delete: List[int]) -> int:
        if not rows_to_delete:
            return 0

        worksheet = self.com_manager.worksheet
        excel_app = self.com_manager.excel_app

        sorted_rows = sorted(rows_to_delete, reverse=True)
        deleted_count = 0

        try:
            if excel_app:
                excel_app.ScreenUpdating = False

            for row in sorted_rows:
                worksheet.Rows(row).Delete()
                deleted_count += 1
                logger.info(f"Đã xóa dòng {row}")

            logger.info(f"Đã xóa tổng cộng {deleted_count} dòng trùng")
            return deleted_count

        except Exception as e:
            logger.error(f"Lỗi khi xóa dòng (đã xóa {deleted_count}/{len(sorted_rows)}): {e}")
            raise RuntimeError(
                f"Lỗi khi xóa dòng trùng (đã xóa {deleted_count}/{len(sorted_rows)}): {str(e)}"
            )
        finally:
            if excel_app:
                excel_app.ScreenUpdating = True
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_duplicate_size_detector.py -v`
Expected: 8 tests PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_duplicate_size_detector.py excel_automation/duplicate_size_detector.py
git commit -m "feat: add DuplicateSizeDetector.delete_rows with bottom-up deletion"
```

---

## Task 3: Export DuplicateSizeDetector

**Files:**
- Modify: `excel_automation/__init__.py`

- [ ] **Step 1: Add import and export**

Trong `excel_automation/__init__.py`, thêm import sau dòng `from excel_automation.excel_com_manager import ExcelCOMManager`:

```python
from excel_automation.duplicate_size_detector import DuplicateSizeDetector
```

Thêm `"DuplicateSizeDetector"` vào list `__all__`, sau `"ExcelCOMManager"`:

```python
    "ExcelCOMManager",
    "DuplicateSizeDetector",
```

- [ ] **Step 2: Run existing tests to ensure no breakage**

Run: `pytest tests/ -v`
Expected: All existing tests PASS

- [ ] **Step 3: Commit**

```bash
git add excel_automation/__init__.py
git commit -m "feat: export DuplicateSizeDetector from excel_automation package"
```

---

## Task 4: DuplicateSizeDialog — UI

**Files:**
- Create: `ui/duplicate_size_dialog.py`

- [ ] **Step 1: Create dialog file**

File: `ui/duplicate_size_dialog.py`

```python
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List
import logging

logger = logging.getLogger(__name__)

BASE_HEIGHT = 160
ROW_HEIGHT = 30
GROUP_HEADER_HEIGHT = 35
MIN_HEIGHT = 200
MAX_HEIGHT_RATIO = 0.7
DIALOG_WIDTH = 450


class DuplicateSizeDialog:

    def __init__(self, parent: tk.Tk, duplicate_sizes: Dict[str, List[int]]):
        self.parent = parent
        self.duplicate_sizes = duplicate_sizes
        self.group_checkboxes: Dict[str, Dict[int, tk.BooleanVar]] = {}
        self.rows_to_delete: List[int] = []

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Phát hiện Size trùng")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_skip)

        self._create_widgets()
        self._calculate_and_set_size()
        self._center_window()

    def _calculate_and_set_size(self) -> None:
        total_rows = sum(len(rows) for rows in self.duplicate_sizes.values())
        total_groups = len(self.duplicate_sizes)

        content_height = (
            BASE_HEIGHT
            + total_groups * GROUP_HEADER_HEIGHT
            + total_rows * ROW_HEIGHT
        )

        screen_height = self.parent.winfo_screenheight()
        max_height = int(screen_height * MAX_HEIGHT_RATIO)

        height = max(MIN_HEIGHT, min(content_height, max_height))
        self.dialog.geometry(f"{DIALOG_WIDTH}x{height}")

    def _center_window(self) -> None:
        self.dialog.update_idletasks()
        w = self.dialog.winfo_width()
        h = self.dialog.winfo_height()
        x = (self.parent.winfo_screenwidth() // 2) - (w // 2)
        y = (self.parent.winfo_screenheight() // 2) - (h // 2)
        self.dialog.geometry(f"+{x}+{y}")

    def _create_widgets(self) -> None:
        header_frame = ttk.Frame(self.dialog)
        header_frame.pack(fill=tk.X, padx=10, pady=10)

        total_sizes = len(self.duplicate_sizes)
        total_rows = sum(len(rows) for rows in self.duplicate_sizes.values())

        ttk.Label(
            header_frame,
            text=f"Phát hiện {total_sizes} size trùng ({total_rows} dòng)",
            font=('Arial', 11, 'bold'),
            foreground='#d35400'
        ).pack(anchor=tk.W)

        ttk.Label(
            header_frame,
            text="Check dòng muốn GIỮ, dòng không check sẽ bị XÓA.",
            font=('Arial', 9),
            foreground='gray'
        ).pack(anchor=tk.W, pady=(5, 0))

        scroll_frame = ttk.Frame(self.dialog)
        scroll_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        canvas = tk.Canvas(scroll_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        )

        from excel_automation.utils import get_size_sort_key
        sorted_sizes = sorted(self.duplicate_sizes.keys(), key=get_size_sort_key)

        for size in sorted_sizes:
            rows = self.duplicate_sizes[size]

            group_frame = ttk.LabelFrame(
                scrollable_frame,
                text=f"Size {size} ({len(rows)} dòng)",
                padding=5
            )
            group_frame.pack(fill=tk.X, padx=5, pady=(5, 0))

            self.group_checkboxes[size] = {}

            for idx, row in enumerate(sorted(rows)):
                var = tk.BooleanVar(value=(idx == 0))
                self.group_checkboxes[size][row] = var

                cb = ttk.Checkbutton(
                    group_frame,
                    text=f"Dòng {row}",
                    variable=var
                )
                cb.pack(anchor=tk.W, padx=10, pady=2)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        action_frame = ttk.Frame(self.dialog)
        action_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Button(
            action_frame,
            text="Xóa dòng trùng",
            command=self._on_delete,
            width=18
        ).pack(side=tk.RIGHT, padx=(5, 0))

        ttk.Button(
            action_frame,
            text="Bỏ qua",
            command=self._on_skip,
            width=12
        ).pack(side=tk.RIGHT)

    def _on_delete(self) -> None:
        for size, row_vars in self.group_checkboxes.items():
            checked_count = sum(1 for var in row_vars.values() if var.get())
            if checked_count == 0:
                messagebox.showwarning(
                    "Cảnh báo",
                    f"Size {size} phải giữ ít nhất 1 dòng!\n"
                    f"Vui lòng check ít nhất 1 dòng cho size {size}.",
                    parent=self.dialog
                )
                return

        rows_to_delete = []
        for size, row_vars in self.group_checkboxes.items():
            for row, var in row_vars.items():
                if not var.get():
                    rows_to_delete.append(row)

        if not rows_to_delete:
            messagebox.showinfo(
                "Thông báo",
                "Tất cả dòng đều được giữ, không có gì để xóa.",
                parent=self.dialog
            )
            self.dialog.destroy()
            return

        confirm = messagebox.askyesno(
            "Xác nhận xóa",
            f"Bạn có chắc muốn XÓA {len(rows_to_delete)} dòng?\n\n"
            f"Dòng sẽ xóa: {', '.join(str(r) for r in sorted(rows_to_delete))}\n\n"
            f"Hành động này không thể hoàn tác!",
            parent=self.dialog
        )

        if confirm:
            self.rows_to_delete = rows_to_delete
            self.dialog.destroy()

    def _on_skip(self) -> None:
        self.rows_to_delete = []
        self.dialog.destroy()

    def get_rows_to_delete(self) -> List[int]:
        return self.rows_to_delete

    def show(self) -> None:
        self.parent.wait_window(self.dialog)
```

- [ ] **Step 2: Verify syntax**

Run: `python -c "import ast; ast.parse(open('ui/duplicate_size_dialog.py').read()); print('OK')"`
Expected: `OK`

- [ ] **Step 3: Commit**

```bash
git add ui/duplicate_size_dialog.py
git commit -m "feat: add DuplicateSizeDialog for selecting duplicate rows to keep/delete"
```

---

## Task 5: Tích hợp vào copy sheet flow

**Files:**
- Modify: `ui/excel_realtime_controller.py` — method `_copy_sheet_continue` (line ~471) và thêm method `_check_and_remove_duplicate_sizes`

- [ ] **Step 1: Add _check_and_remove_duplicate_sizes method**

Thêm method mới vào class `ExcelRealtimeController`, ngay sau method `_copy_sheet_retry` (sau dòng 516):

```python
    def _check_and_remove_duplicate_sizes(self) -> None:
        try:
            from excel_automation.duplicate_size_detector import DuplicateSizeDetector
            from ui.duplicate_size_dialog import DuplicateSizeDialog

            detector = DuplicateSizeDetector(self.com_manager)
            duplicates = detector.detect_duplicates()

            if not duplicates:
                logger.info("Không có size trùng sau copy sheet")
                return

            total_dup_rows = sum(len(rows) for rows in duplicates.values())
            logger.info(f"Phát hiện {len(duplicates)} size trùng ({total_dup_rows} dòng)")

            dialog = DuplicateSizeDialog(self.root, duplicates)
            dialog.show()

            rows_to_delete = dialog.get_rows_to_delete()

            if not rows_to_delete:
                logger.info("User bỏ qua xóa size trùng")
                return

            deleted = detector.delete_rows(rows_to_delete)

            self._scan_sizes()
            self._deselect_all_sizes()

            self.status_label.config(
                text=f"Đã xóa {deleted} dòng size trùng"
            )
            logger.info(f"Đã xóa {deleted} dòng size trùng sau copy sheet")

        except Exception as e:
            logger.error(f"Lỗi khi xử lý size trùng: {e}")
            messagebox.showerror(
                "Lỗi",
                f"Lỗi khi xóa dòng trùng:\n{str(e)}"
            )
```

- [ ] **Step 2: Modify _copy_sheet_continue to call duplicate detection**

Trong method `_copy_sheet_continue`, thêm gọi `_check_and_remove_duplicate_sizes()` sau `progress.complete_step(2)` và trước `progress.start_step(3)`.

Tìm đoạn code này trong `_copy_sheet_continue` (khoảng dòng 480-485):

```python
            progress.start_step(2)
            self._scan_sizes()
            self._deselect_all_sizes()
            progress.complete_step(2)

            progress.start_step(3)
```

Thay bằng:

```python
            progress.start_step(2)
            self._scan_sizes()
            self._deselect_all_sizes()
            progress.complete_step(2)

            progress.dialog.withdraw()
            self._check_and_remove_duplicate_sizes()
            progress.dialog.deiconify()
            progress.dialog.grab_set()

            progress.start_step(3)
```

`progress.dialog.withdraw()` ẩn progress dialog tạm thời để DuplicateSizeDialog có thể grab focus. Sau khi xong thì `deiconify()` + `grab_set()` để progress dialog tiếp tục.

- [ ] **Step 3: Test manually**

1. Mở app: `python excel_realtime_controller.py`
2. Mở một file Excel packing list có dữ liệu size trong cột F
3. Thử copy sheet (nút "Copy Sheet")
4. Nếu không có size trùng → flow bình thường, không dialog nào hiện
5. Nếu có size trùng → dialog "Phát hiện Size trùng" hiện ra
6. Check dòng muốn giữ → nhấn "Xóa dòng trùng" → xác nhận → dòng bị xóa
7. Kiểm tra Excel: dòng đã bị xóa, các dòng dồn lên

- [ ] **Step 4: Commit**

```bash
git add ui/excel_realtime_controller.py
git commit -m "feat: integrate duplicate size detection into copy sheet flow"
```

---

## Task 6: Final verification

- [ ] **Step 1: Run all tests**

Run: `pytest tests/ -v`
Expected: All tests PASS

- [ ] **Step 2: Commit nếu có thay đổi cuối**

Chỉ commit nếu có fix gì thêm trong quá trình test.
