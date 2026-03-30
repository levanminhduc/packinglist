# Duplicate Size Detection — Quét & Xóa Size Trùng Sau Khi Copy Sheet

## Bối cảnh

Khi copy sheet trong packing list, cột F (size column) có thể chứa nhiều dòng cùng một size (ví dụ size 060 xuất hiện ở dòng 19, 25, 31). Hiện tại `scan_sizes()` dùng `set()` nên tự loại trùng — user không biết có dòng trùng. Cần tính năng tự phát hiện và cho phép xóa dòng trùng.

## Yêu cầu

1. Sau khi copy sheet xong, tự động quét cột F tìm size trùng (size xuất hiện >= 2 dòng)
2. Nếu có trùng, mở dialog cho user chọn dòng giữ/xóa
3. Dòng được check = giữ, dòng không check = xóa hẳn (delete row, dồn lên)
4. Nếu không có trùng, bỏ qua — không hiện dialog

## Thiết kế

### Luồng hoạt động

1. `_copy_sheet_continue()` hoàn tất các bước copy sheet hiện tại
2. Gọi `ExcelCOMManager.get_size_row_mapping()` để lấy mapping size → list rows
3. Lọc ra `duplicate_sizes = {size: rows for size, rows in mapping.items() if len(rows) >= 2}`
4. Nếu `duplicate_sizes` rỗng → bỏ qua
5. Nếu có → mở `DuplicateSizeDialog(parent, duplicate_sizes)`
6. User chọn dòng giữ → nhấn "Xóa dòng trùng"
7. Xác nhận trước khi xóa (messagebox.askyesno)
8. Xóa dòng từ lớn → nhỏ (bottom-up) để tránh lệch index
9. Quét lại sizes sau khi xóa

### Tích hợp vào copy sheet flow

Vị trí: trong `_copy_sheet_continue()`, sau step 2 (scan sizes) và trước step 3 (show all rows). Thêm một bước mới gọi `_check_and_remove_duplicate_sizes()`.

### Business logic — `DuplicateSizeDetector`

File: `excel_automation/duplicate_size_detector.py`

Class: `DuplicateSizeDetector`
- Input: `ExcelCOMManager` instance
- `detect_duplicates(column, start_row, end_row) -> Dict[str, List[int]]`: Dùng lại logic từ `ExcelCOMManager.scan_sizes()` (COM Range read + `normalize_size_value`) để build size→rows mapping, lọc trả về chỉ những size có >= 2 dòng
- `delete_rows(worksheet, excel_app, rows_to_delete: List[int]) -> int`: Xóa danh sách dòng (bottom-up), trả về số dòng đã xóa. Wrap `ScreenUpdating = False/True`

### Dialog — `DuplicateSizeDialog`

File: `ui/duplicate_size_dialog.py`

Layout:
```
┌─ Phát hiện Size trùng ──────────────────────┐
│                                              │
│  ⚠ Phát hiện X size trùng trong cột F       │
│  Check dòng muốn GIỮ, dòng không check      │
│  sẽ bị XÓA.                                 │
│                                              │
│  ── Size 060 (3 dòng) ──────────────────     │
│  ☑ Dòng 19                                   │
│  ☐ Dòng 25                                   │
│  ☐ Dòng 31                                   │
│                                              │
│  ── Size 044 (2 dòng) ──────────────────     │
│  ☑ Dòng 20                                   │
│  ☐ Dòng 38                                   │
│                                              │
│         [Bỏ qua]  [Xóa dòng trùng]          │
└──────────────────────────────────────────────┘
```

Hành vi:
- Mặc định: dòng đầu tiên mỗi nhóm được check (giữ), còn lại không check (xóa)
- Tự co giãn chiều cao theo tổng số dòng trùng: `base_height + (total_duplicate_rows * row_height)`, min 200px, max 70% màn hình
- Nếu vượt max → có scrollbar
- Nút "Bỏ qua": đóng dialog, không xóa gì, trả về empty list
- Nút "Xóa dòng trùng": thu thập dòng không check, hiện xác nhận, xóa nếu đồng ý
- Validation: mỗi nhóm size phải có ít nhất 1 dòng được check (không cho xóa hết tất cả dòng của 1 size)

### Xóa dòng

- Thu thập tất cả dòng không check từ tất cả nhóm
- Sort descending (dòng lớn nhất trước)
- Loop: `worksheet.Rows(row).Delete()` cho từng dòng
- Wrap `ScreenUpdating = False / True`
- Log mỗi dòng đã xóa

### Error handling

- Nếu xóa dòng lỗi giữa chừng → log error, hiện thông báo với số dòng đã xóa thành công
- Dialog bị đóng (X) → giống nhấn "Bỏ qua"

## Các file cần thay đổi

1. **Tạo mới**: `excel_automation/duplicate_size_detector.py` — Business logic detect + delete
2. **Tạo mới**: `ui/duplicate_size_dialog.py` — Dialog UI
3. **Sửa**: `ui/excel_realtime_controller.py` — Thêm gọi detect sau copy sheet trong `_copy_sheet_continue()`
4. **Sửa**: `excel_automation/__init__.py` — Export class mới

## Ngoài phạm vi

- Không thêm nút/menu riêng (chỉ tự động sau copy sheet)
- Không hiển thị thông tin cột khác ngoài số dòng
- Không undo sau khi xóa
