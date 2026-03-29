# Auto Ghi Excel + Ẩn Dòng Sau Import PO — Design Spec

## Mục tiêu

Sau khi Import PO từ PDF thành công, tự động thực hiện 2 bước tiếp theo mà hiện tại user phải bấm tay:
1. **Ghi Sizes & Quantities vào Excel** (nút "💾 Ghi vào Excel")
2. **Ẩn dòng không chọn** (nút "👁️ Ẩn dòng ngay")

Chỉ hiện 1 messagebox tổng hợp "Import thành công" cuối cùng.

## Phạm vi thay đổi

| Action | File | Thay đổi |
|--------|------|----------|
| Modify | `ui/pdf_import_dialog.py` | Thêm 2 steps vào `ImportProgressDialog.STEPS` và `STEP_WEIGHTS` |
| Modify | `ui/excel_realtime_controller.py` | Thêm logic ghi + ẩn trong `run_write_steps()` của `_execute_import()` |

## Chi tiết

### 1. Mở rộng ImportProgressDialog

**Trước (7 bước, index 0-6):**
```
0: Đọc file PDF
1: Trích xuất dữ liệu PO, Color, Sizes
2: Scan sizes từ Excel
3: Ghi PO vào Excel
4: Ghi Color Code vào Excel
5: Cập nhật Sizes & Quantities
6: Hoàn tất
```

**Sau (9 bước, index 0-8):**
```
0: Đọc file PDF
1: Trích xuất dữ liệu PO, Color, Sizes
2: Scan sizes từ Excel
3: Ghi PO vào Excel
4: Ghi Color Code vào Excel
5: Cập nhật Sizes & Quantities     ← giữ nguyên (tick checkbox + điền entry UI)
6: Ghi Sizes & Quantities vào Excel ← MỚI
7: Ẩn dòng không chọn               ← MỚI
8: Hoàn tất
```

**STEP_WEIGHTS** cần điều chỉnh lại tổng 100:
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

Dialog height cần tăng từ 380 lên ~430 để chứa 9 steps.

### 2. Thêm logic trong _execute_import

Trong `run_write_steps()`, sau step 5 (`_apply_imported_sizes`):

**Step 6 — Ghi Sizes & Quantities vào Excel:**
- `selected_sizes = [size for size in size_quantities.keys() if size in self.checkboxes]`
- Tạo `SizeQuantityDisplayManager(self.config)`
- Lấy `current_quantities` từ Excel
- Gọi `display_manager.write_quantities_to_excel()`
- Hủy auto-save timer (vì `_apply_imported_sizes` đã gọi `_reset_auto_save_timer()`, nhưng ta đã ghi rồi nên không cần đợi 10s nữa)

**Step 7 — Ẩn dòng không chọn:**
- Gọi `self.com_manager.hide_rows_realtime(selected_sizes)`

**Step 8 — Hoàn tất** (dịch từ index 6 → 8):
- Giữ nguyên: `_update_po_color_display()`, `_update_box_count_display()`

### 3. Messagebox cuối

Gộp tất cả info thành 1 thông báo:
```
Đã import PO từ PDF:

PO: {po}
Color: {color}
Sizes: {len(size_quantities)}
Total Qty: {total:,}
Đã ghi {written_count} cells vào Excel
Đã ẩn {hidden_count} dòng
```

### 4. Luồng dữ liệu

`_execute_import(po, color, size_quantities)` đã có sẵn `size_quantities: Dict[str, int]`. Không cần đọc lại từ UI.

`selected_sizes` lấy từ `size_quantities.keys()` lọc qua `self.checkboxes` (đảm bảo chỉ ghi size có trong Excel).

### 5. Hủy auto-save timer

`_apply_imported_sizes()` gọi `_reset_auto_save_timer()` ở cuối, bắt đầu countdown 10s. Sau khi ghi trực tiếp ở step 6, cần hủy timer này bằng cách:
```python
if self._auto_save_timer_id is not None:
    self.root.after_cancel(self._auto_save_timer_id)
    self._auto_save_timer_id = None
    self._auto_save_pending = False
```

### 6. Error handling

Mỗi step mới đều nằm trong `run_write_steps()` có sẵn cơ chế try/except + `progress.show_error(step, str(e), lambda: run_write_steps(step))` để retry từ bước lỗi. Không cần thêm error handling mới.

## Không thay đổi

- Các nút "💾 Ghi vào Excel" và "👁️ Ẩn dòng ngay" vẫn hoạt động độc lập khi user bấm tay
- Auto-save 10s vẫn hoạt động bình thường cho flow không qua import
- Flow parse PDF → preview dialog giữ nguyên
