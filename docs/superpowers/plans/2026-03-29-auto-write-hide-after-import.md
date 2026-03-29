# Auto Ghi Excel + Ẩn Dòng Sau Import PO — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Sau khi Import PO từ PDF thành công, tự động Ghi Sizes & Quantities vào Excel rồi Ẩn dòng không chọn — thay vì user phải bấm tay 2 nút riêng.

**Architecture:** Mở rộng `ImportProgressDialog.STEPS` thêm 2 bước mới (index 6, 7), dịch "Hoàn tất" sang index 8. Trong `_execute_import().run_write_steps()`, thêm logic gọi `SizeQuantityDisplayManager.write_quantities_to_excel()` và `com_manager.hide_rows_realtime()`. Messagebox cuối gộp tất cả info.

**Tech Stack:** Python 3, tkinter, win32com (COM), pytest

**Spec:** `docs/superpowers/specs/2026-03-29-auto-write-hide-after-import-design.md`

---

## File Structure

| Action | File | Responsibility |
|--------|------|----------------|
| Modify | `ui/pdf_import_dialog.py` | Thêm 2 entries vào `STEPS`, `STEP_WEIGHTS`, tăng dialog height |
| Modify | `ui/excel_realtime_controller.py` | Thêm step 6 (ghi), step 7 (ẩn), sửa step index "Hoàn tất", sửa messagebox |
| Create | `tests/test_auto_write_hide_after_import.py` | Unit tests cho logic mới |

---

### Task 1: Mở rộng ImportProgressDialog — thêm 2 steps mới

**Files:**
- Modify: `ui/pdf_import_dialog.py:276-286` (STEPS + STEP_WEIGHTS)
- Modify: `ui/pdf_import_dialog.py:296` (dialog height)

- [ ] **Step 1: Sửa STEPS list — thêm 2 entries mới trước "Hoàn tất"**

Trong `ui/pdf_import_dialog.py`, thay thế `STEPS` và `STEP_WEIGHTS`:

```python
    STEPS = [
        "Đọc file PDF",
        "Trích xuất dữ liệu PO, Color, Sizes",
        "Scan sizes từ Excel",
        "Ghi PO vào Excel",
        "Ghi Color Code vào Excel",
        "Cập nhật Sizes & Quantities",
        "Ghi Sizes & Quantities vào Excel",
        "Ẩn dòng không chọn",
        "Hoàn tất",
    ]

    STEP_WEIGHTS = [15, 10, 10, 12, 12, 10, 13, 13, 5]
```

- [ ] **Step 2: Tăng dialog height từ 380 lên 430**

Trong `ui/pdf_import_dialog.py`, thay:

```python
        self.dialog.geometry("420x380")
```

thành:

```python
        self.dialog.geometry("420x430")
```

- [ ] **Step 3: Commit**

```bash
git add ui/pdf_import_dialog.py
git commit -m "feat: add 2 new progress steps for auto write+hide after import"
```

---

### Task 2: Thêm step 6 — Ghi Sizes & Quantities vào Excel

**Files:**
- Modify: `ui/excel_realtime_controller.py:1340-1383` (method `_execute_import`, inner function `run_write_steps`)

- [ ] **Step 1: Thêm step 6 vào run_write_steps()**

Trong `ui/excel_realtime_controller.py`, method `_execute_import`, inner function `run_write_steps()`, thay thế toàn bộ block từ dòng `progress.start_step(6)` đến `progress.finish()` (dòng 1359-1376).

Block cũ:

```python
                progress.start_step(6)
                self._update_po_color_display()
                self._update_box_count_display()
                progress.complete_step(6)

                progress.finish()

                self.status_label.config(
                    text=f"Import thành công: PO={po}, Color={color}, {len(size_quantities)} sizes"
                )
                messagebox.showinfo(
                    "Thành Công",
                    f"Đã import PO từ PDF:\n\n"
                    f"PO: {po}\n"
                    f"Color: {color}\n"
                    f"Sizes: {len(size_quantities)}\n"
                    f"Total Qty: {sum(size_quantities.values()):,}"
                )
```

Block mới:

```python
                if start_from <= 6:
                    progress.start_step(6)
                    selected_sizes = [
                        size for size in size_quantities.keys()
                        if size in self.checkboxes
                    ]
                    display_manager = SizeQuantityDisplayManager(self.config)
                    current_quantities = display_manager.get_current_quantities(
                        self.com_manager.worksheet,
                        selected_sizes,
                        self.config.get_column()
                    )
                    written_count = display_manager.write_quantities_to_excel(
                        self.com_manager.excel_app,
                        self.com_manager.worksheet,
                        selected_sizes,
                        size_quantities,
                        current_quantities,
                        self.config.get_column()
                    )
                    if self._auto_save_timer_id is not None:
                        self.root.after_cancel(self._auto_save_timer_id)
                        self._auto_save_timer_id = None
                        self._auto_save_pending = False
                    progress.complete_step(6)

                if start_from <= 7:
                    progress.start_step(7)
                    if not selected_sizes:
                        selected_sizes = [
                            size for size in size_quantities.keys()
                            if size in self.checkboxes
                        ]
                    hidden_count = self.com_manager.hide_rows_realtime(selected_sizes)
                    progress.complete_step(7)

                progress.start_step(8)
                self._update_po_color_display()
                self._update_box_count_display()
                progress.complete_step(8)

                progress.finish()

                self.status_label.config(
                    text=f"Import thành công: PO={po}, Color={color}, {len(size_quantities)} sizes"
                )
                messagebox.showinfo(
                    "Thành Công",
                    f"Đã import PO từ PDF:\n\n"
                    f"PO: {po}\n"
                    f"Color: {color}\n"
                    f"Sizes: {len(size_quantities)}\n"
                    f"Total Qty: {sum(size_quantities.values()):,}\n"
                    f"Đã ghi {written_count} cells vào Excel\n"
                    f"Đã ẩn {hidden_count} dòng"
                )
```

Lưu ý: biến `written_count` và `hidden_count` cần được khai báo trước vòng `if start_from` để tránh `UnboundLocalError` khi retry từ step > 6. Thêm ở đầu `run_write_steps()`, ngay sau `def run_write_steps(start_from: int = 3):`:

```python
            nonlocal written_count, hidden_count
```

Và khai báo 2 biến này trước `def run_write_steps`:

```python
        written_count = 0
        hidden_count = 0
```

- [ ] **Step 2: Commit**

```bash
git add ui/excel_realtime_controller.py
git commit -m "feat: auto write quantities + hide rows after PO import"
```

---

### Task 3: Unit tests

**Files:**
- Create: `tests/test_auto_write_hide_after_import.py`

- [ ] **Step 1: Viết tests**

Tạo file `tests/test_auto_write_hide_after_import.py`:

```python
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
```

- [ ] **Step 2: Chạy tests**

```bash
pytest tests/test_auto_write_hide_after_import.py -v
```

Expected: tất cả PASS.

- [ ] **Step 3: Commit**

```bash
git add tests/test_auto_write_hide_after_import.py
git commit -m "test: add tests for auto write+hide progress steps"
```

---

### Task 4: Manual smoke test

- [ ] **Step 1: Chạy app và kiểm tra flow**

```bash
python excel_realtime_controller.py
```

Kiểm tra:
1. Mở file Excel, chọn sheet
2. Bấm "📄 Import PO từ PDF" → chọn file PDF
3. Preview dialog hiện → bấm "Xác nhận"
4. Progress dialog phải hiện 9 bước, bao gồm "Ghi Sizes & Quantities vào Excel" và "Ẩn dòng không chọn"
5. Sau khi xong, chỉ hiện 1 messagebox "Import thành công" với thông tin ghi cells + ẩn dòng
6. Excel phải có quantities đã ghi VÀ dòng không chọn đã bị ẩn

Kiểm tra thêm:
- Nút "💾 Ghi vào Excel" vẫn hoạt động bình thường khi bấm tay
- Nút "👁️ Ẩn dòng ngay" vẫn hoạt động bình thường khi bấm tay

- [ ] **Step 2: Chạy full test suite**

```bash
pytest tests/ -v
```

Expected: tất cả PASS, không có regression.

- [ ] **Step 3: Commit (nếu có fix gì)**

```bash
git add -A
git commit -m "fix: address issues found during smoke test"
```
