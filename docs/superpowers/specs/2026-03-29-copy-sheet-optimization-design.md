# Copy Sheet Performance Optimization

## Problem

Flow Copy Sheet gây đơ cả app tkinter và Excel cùng lúc (not responding) trong vài giây. Root cause: ~1,500 COM calls riêng lẻ chạy synchronous trên UI thread.

### Phân tích bottleneck

Khi user nhấn Copy Sheet, flow thực hiện tuần tự:

| Bước | Method | COM calls | Ghi chú |
|---|---|---|---|
| 1 | `copy_sheet()` | 1 | `source_sheet.Copy()` — chấp nhận được |
| 2 | `rename_sheet()` | 1 | Dialog + COM rename |
| 3 | `switch_sheet()` | 1 | Activate sheet mới |
| 4 | `clear_quantity_columns()` | ~1,353 | Loop 41 rows × 33 cols, đọc + ghi từng cell |
| 5 | `scan_sizes()` | ~41 | Loop từng cell đọc size |
| 6 | `_load_quantities_from_excel()` | ~41 | Loop từng cell đọc quantity |
| 7 | `show_all_rows()` | ~41 | Loop từng row set Hidden = False |
| 8 | `_update_po_color_display()` | ~3 | Đọc PO + Color |
| **Tổng** | | **~1,481** | |

Mỗi COM call = 1 IPC round-trip giữa Python process và Excel process.

### Constraints

- Sheet nhỏ: dưới 50 dòng data (row 19 trở xuống)
- Chỉ Copy Sheet flow bị đơ, các thao tác khác chấp nhận được
- Windows-only, COM automation qua `win32com.client`

## Solution

Kết hợp 2 lớp tối ưu:

1. **Bulk Range Operations** — giảm COM calls từ ~1,500 xuống ~5-10
2. **Progress Dialog** — hiển thị tiến trình realtime để user biết app đang xử lý

### 1. Bulk Range Operations trong ExcelCOMManager

#### `clear_quantity_columns()`

Hiện tại: loop từng cell đọc rồi ghi None

```python
for row in range(start_row, end_row + 1):
    for col in range(start_col, end_col + 1):
        cell_value = self.worksheet.Cells(row, col).Value
        if cell_value is not None:
            self.worksheet.Cells(row, col).Value = None
```

Sau tối ưu: 1 lệnh ClearContents cho toàn bộ range

```python
start_col_letter = self._number_to_column_letter(start_col)
end_col_letter = self._number_to_column_letter(end_col)
range_str = f"{start_col_letter}{start_row}:{end_col_letter}{end_row}"
target_range = self.worksheet.Range(range_str)
cleared_count = self.excel_app.WorksheetFunction.CountA(target_range)
target_range.ClearContents
```

Dùng `CountA` trước để đếm cells có data (trả về count chính xác), rồi `ClearContents` xóa 1 lần. Tổng: 2 COM calls thay vì ~1,353.

#### `show_all_rows()`

Hiện tại: loop từng row

```python
for row in range(start_row, end_row + 1):
    self.worksheet.Rows(row).Hidden = False
```

Sau tối ưu: 1 lệnh cho toàn bộ range

```python
range_str = f"{start_row}:{end_row}"
self.worksheet.Range(range_str).EntireRow.Hidden = False
```

#### `scan_sizes()`

Hiện tại: loop từng cell đọc value

```python
for row in range(start_row, end_row + 1):
    cell_value = self.worksheet.Cells(row, col_num).Value
```

Sau tối ưu: đọc cả range 1 lần

```python
range_str = f"{col_letter}{start_row}:{col_letter}{end_row}"
values = self.worksheet.Range(range_str).Value
```

`Range.Value` trả về tuple of tuples `((val1,), (val2,), ...)`. Xử lý normalize trong Python (không cần COM).

Lưu ý: `_fix_decimal_cell()` vẫn ghi từng cell lẻ khi gặp số thập phân, nhưng trường hợp này hiếm nên chấp nhận được.

#### Helper method mới

Thêm `_number_to_column_letter(col_num: int) -> str` để convert số cột sang letter (7 -> "G", 39 -> "AM"). Ngược lại với `_column_letter_to_number()` đã có.

### 2. CopySheetProgressDialog

#### File mới: `ui/copy_sheet_progress_dialog.py`

Reuse pattern từ `ImportProgressDialog` trong `ui/pdf_import_dialog.py`:

```python
class CopySheetProgressDialog:
    STEPS = [
        "Copy sheet",
        "Xóa số lượng cũ",
        "Quét sizes",
        "Cập nhật giao diện",
        "Hoàn tất",
    ]

    STEP_WEIGHTS = [30, 25, 20, 20, 5]
```

Cấu trúc dialog:
- Title: "Đang Copy Sheet..."
- Size: 420×350, không resize
- Modal: `transient(parent)`, `grab_set()`, block `WM_DELETE_WINDOW`
- Progress bar determinate + percent label
- Step list với icons: `⬚` pending, `🔄` in progress, `✅` done, `❌` error
- Khi lỗi: hiện error message + nút "Thử lại" + nút "Đóng"
- Khi hoàn tất: auto close sau 1 giây

Methods (giống ImportProgressDialog):
- `start_step(step_index)` — mark step đang xử lý, cập nhật progress bar
- `complete_step(step_index)` — mark step hoàn thành
- `finish()` — set 100%, auto close sau 1s
- `show_error(step_index, error_msg, retry_callback)` — hiện lỗi + nút retry

### 3. Flow mới của `_copy_sheet()` trong UI

```
1. Step 0 - "Copy sheet":
   - Tạo CopySheetProgressDialog
   - progress.start_step(0)
   - com_manager.copy_sheet()
   - progress.complete_step(0)
   - progress.dialog.withdraw() (ẩn tạm)
2. Hiện SheetRenameDialog để user đặt tên
   - Nếu user hủy → đóng progress, reload sheets, return
3. Rename sheet nếu cần, switch_sheet()
   - progress.dialog.deiconify() (hiện lại progress)
4. Step 1 - "Xóa số lượng cũ":
   - progress.start_step(1)
   - com_manager.clear_quantity_columns() (bulk)
   - progress.complete_step(1)
5. Step 2 - "Quét sizes":
   - progress.start_step(2)
   - com_manager.scan_sizes() (bulk)
   - Rebuild UI checkboxes
   - progress.complete_step(2)
6. Step 3 - "Cập nhật giao diện":
   - progress.start_step(3)
   - show_all_rows() (bulk)
   - _update_po_color_display()
   - _highlight_update_buttons()
   - _start_auto_refresh_sizes()
   - Update sheet combobox
   - progress.complete_step(3)
7. Step 4 - "Hoàn tất":
   - progress.finish() → auto close sau 1s
   - Cập nhật status_label
```

Progress dialog dùng `withdraw()`/`deiconify()` để ẩn/hiện — giữ nguyên 1 instance xuyên suốt flow, không tạo mới.

### 4. Error Handling

Mỗi step trong flow được wrap try/except. Khi lỗi:
- `progress.show_error()` hiện thông báo + nút Thử lại
- Retry callback chạy lại từ step bị lỗi (không chạy lại từ đầu)
- ScreenUpdating luôn được restore về True trong finally block

### 5. Files thay đổi

| File | Loại | Thay đổi |
|---|---|---|
| `excel_automation/excel_com_manager.py` | Sửa | Refactor `clear_quantity_columns()`, `show_all_rows()`, `scan_sizes()` dùng bulk Range. Thêm `_number_to_column_letter()` |
| `ui/copy_sheet_progress_dialog.py` | Mới | `CopySheetProgressDialog` class |
| `ui/excel_realtime_controller.py` | Sửa | Refactor `_copy_sheet()` dùng progress dialog + gọi bulk methods |

### 6. Không thay đổi

- Các method khác trong `ExcelCOMManager` (hide_rows_realtime, v.v.) — user xác nhận chỉ Copy Sheet bị lag
- `ImportProgressDialog` — giữ nguyên, không refactor thành base class (YAGNI)
- Flow rename dialog — vẫn xen giữa step 0 và step 1
- `_fix_decimal_cell()` — vẫn ghi từng cell lẻ khi cần
