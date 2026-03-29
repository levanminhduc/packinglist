# Box List Export Optimization

## Vấn đề

Khi xuất danh sách thùng, UI bị đóng băng, Excel nhấp nháy, và thời gian xử lý lâu. Nguyên nhân gốc: hàng trăm COM calls đọc/ghi cell-by-cell, chạy đồng bộ trên main thread mà không có progress indicator.

## Giải pháp

Tối ưu COM batch read/write để giảm thời gian xử lý thực tế + thêm progress dialog để user thấy tiến trình.

## Phạm vi thay đổi

- `excel_automation/box_list_export_manager.py` — tối ưu `read_box_ranges()` và `paste_and_format_to_excel()`
- `ui/box_list_export_progress_dialog.py` — file mới, progress dialog
- `ui/excel_realtime_controller.py` — refactor `_export_box_list()` dùng progress dialog

Output (format trên sheet mới) giữ nguyên 100%.

## 1. Tối ưu batch read — `read_box_ranges()`

### Hiện tại

Đọc cell-by-cell: mỗi size × mỗi cột = 3 COM calls (quantity, box_start, box_end). 10 sizes × 32 cột = ~960 COM calls.

`find_last_data_row()` cũng đọc từng cell một.

### Thiết kế mới

Đọc theo range — mỗi lần đọc cả hàng/cột trong 1 COM call:

1. Đọc cột size bằng `Range(F19:F{end}).Value` → mapping `size → row` trong 1 COM call
2. Đọc hàng `box_start_row` bằng `Range(Cells(row, 7), Cells(row, scan_end)).Value` → tất cả box_start trong 1 COM call
3. Đọc hàng `box_end_row` tương tự → 1 COM call
4. Với mỗi size, đọc hàng data bằng `Range(Cells(size_row, 7), Cells(size_row, scan_end)).Value` → 1 COM call per size

Kết quả: ~960 COM calls → **3 + N calls** (N = số size, thường 5-15). Giảm 50-100x.

Logic xử lý (parse giá trị, validate box_start <= box_end, skip null) giữ nguyên — chỉ thay nguồn dữ liệu từ COM sang tuple 2D đã đọc sẵn.

## 2. Tối ưu batch write — `paste_and_format_to_excel()`

### Hiện tại

Ghi từng cell: mỗi dòng = 3 COM calls (Value, HorizontalAlignment, Font.Bold). 200 dòng = ~600 COM calls.

### Thiết kế mới

Chia thành 2 giai đoạn:

**Giai đoạn 1 — Batch write data:**
- Gom tất cả giá trị vào mảng 2D (list of lists)
- Gán toàn bộ bằng `Range(start, end).Value = array` — 1 COM call per cột

**Giai đoạn 2 — Batch format:**
- Căn giữa toàn bộ vùng data: `Range(toàn_bộ).HorizontalAlignment = -4108` → 1 COM call
- Thu thập các dòng "SIZE ..." → gom row numbers, bold bằng `Union()` để gộp nhiều cell thành 1 range rồi bold 1 lần
- Header (dòng 1): format riêng (Bold, Size 20, căn trái) — 3 COM calls cố định

Kết quả: ~600 COM calls → **5-10 COM calls**. Output giữ nguyên format.

## 3. BoxListExportProgressDialog

### File mới: `ui/box_list_export_progress_dialog.py`

Pattern giống y hệt `CopySheetProgressDialog` và `ImportProgressDialog`.

### Steps và weights

| Step | Tên | Weight |
|------|-----|--------|
| 0 | Đọc dữ liệu thùng từ Excel | 30 |
| 1 | Phân tích & gộp sizes | 15 |
| 2 | Tạo sheet mới | 15 |
| 3 | Ghi danh sách thùng vào sheet | 25 |
| 4 | Copy vào clipboard | 10 |
| 5 | Hoàn tất | 5 |

### Behavior

- Modal dialog: `transient(parent)`, `grab_set()`, block `WM_DELETE_WINDOW`
- Size: 420×350, không resize
- Progress bar determinate + percent label
- Step list với icons: `⬚` pending, `🔄` in progress, `✅` done, `❌` error
- Khi lỗi: hiện error message + nút "Thử lại" + nút "Đóng"
- Khi hoàn tất: auto close sau 1 giây
- `dialog.update()` ở mỗi step transition để UI responsive

### Methods (giống pattern hiện có)

- `start_step(step_index)` — mark step đang xử lý, cập nhật progress bar
- `complete_step(step_index)` — mark step hoàn thành
- `finish()` — set 100%, auto close sau 1s
- `show_error(step_index, error_msg, retry_callback)` — hiện lỗi + nút retry
- `close()` — destroy dialog

## 4. Flow mới của `_export_box_list()`

```
Bấm nút xuất
  → Validate (có COM manager? có size nào được chọn?)
  → Mở BoxListExportProgressDialog
  → Step 0: read_box_ranges (batch read)
  → Step 1: detect_combined_sizes + generate text + split columns
  → Step 2: create_new_sheet
  → Step 3: paste_and_format_to_excel (batch write)
  → Step 4: copy_to_clipboard
  → Step 5: finish → auto close → messagebox thành công
```

Retry: khi lỗi ở step N, retry chạy lại từ step N (giống `_copy_sheet_continue` pattern).

`ScreenUpdating = False` vẫn được wrap ở đầu `export_box_list()` và restore ở `finally`.
