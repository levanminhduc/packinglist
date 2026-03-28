# Fix PDFImportDialog UI Consistency

## Problem

`PDFImportDialog` trong `ui/pdf_import_dialog.py` có 4 vấn đề UI không nhất quán với các dialog khác trong dự án:

1. **Font sai**: Header bảng dùng `("", 9, "bold")` (font rỗng) thay vì `('Arial', 9, 'bold')` như chuẩn dự án
2. **Không cuộn chuột được**: Canvas chứa danh sách size thiếu bind `<MouseWheel>` — chỉ kéo scrollbar bằng tay
3. **Thiếu DEFAULT_CONFIG**: `DialogConfigManager.DEFAULT_CONFIG` không có entry `pdf_import` → fallback `(400, 300)` lần đầu mở
4. **Không lưu vị trí cửa sổ**: Chỉ lưu width/height, mỗi lần mở lại luôn căn giữa màn hình

## Solution

### 1. Font header bảng

Sửa dòng header trong `_create_size_table()`:

- `("", 9, "bold")` → `('Arial', 9, 'bold')`
- Các font `Consolas` cho data entry/label giữ nguyên (đúng pattern monospace cho dữ liệu)

### 2. MouseWheel scrolling

Thêm bind mousewheel cho Canvas trong `_create_size_table()`, copy pattern từ `SizeQuantityInputDialog`:

- `<MouseWheel>` cho Windows/macOS
- `<Button-4>`, `<Button-5>` cho Linux
- Bind lên cả canvas và scrollable frame con

Thêm 2 method:
- `_on_canvas_mousewheel(event)` — xử lý Windows/macOS
- `_on_canvas_mousewheel_linux(event, direction)` — xử lý Linux

### 3. DEFAULT_CONFIG entry

Thêm vào `DialogConfigManager.DEFAULT_CONFIG["dialogs"]`:

```python
"pdf_import": {"width": 600, "height": 550}
```

### 4. Lưu vị trí cửa sổ (chỉ PDFImportDialog)

Thêm 2 method mới vào `DialogConfigManager`:

- `save_dialog_geometry(dialog_name, width, height, x, y)` — lưu đầy đủ 4 giá trị vào config JSON
- `get_dialog_geometry(dialog_name)` → `Tuple[int, int, Optional[int], Optional[int]]` — đọc đầy đủ, x/y fallback `None`

Cập nhật `PDFImportDialog`:

- `__init__`: gọi `get_dialog_geometry('pdf_import')`, nếu có x/y thì set geometry trực tiếp, nếu không thì `_center_window()`
- `_save_size_and_close()`: đổi sang gọi `save_dialog_geometry('pdf_import', w, h, x, y)`

Các dialog khác không bị ảnh hưởng — vẫn dùng `save_dialog_size()` / `get_dialog_size()` cũ.

## Files Changed

| File | Change |
|---|---|
| `ui/pdf_import_dialog.py` | Fix font, add mousewheel, save/restore geometry |
| `excel_automation/dialog_config_manager.py` | Add `pdf_import` to DEFAULT_CONFIG, add `save_dialog_geometry()` + `get_dialog_geometry()` |

## Testing

- Mở PDF import dialog → font header phải là Arial
- Danh sách size nhiều dòng → cuộn chuột hoạt động
- Đóng và mở lại dialog → kích thước và vị trí được nhớ
- Lần đầu mở (chưa có config) → kích thước mặc định 600x550, căn giữa màn hình
