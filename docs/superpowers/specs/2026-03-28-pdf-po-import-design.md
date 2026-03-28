# Import PO từ PDF — Design Spec

## Tổng quan

Tính năng cho phép user chọn file PDF Purchase Order, app tự động trích xuất 4 thông tin (PO, Color, Sizes, Quantities), hiển thị preview cho user review/sửa, sau đó ghi thẳng vào Excel đang mở.

## Quyết định thiết kế

| Quyết định | Kết quả |
|---|---|
| PDF format | Chủ yếu 1 format chính (Hultafors Group), thiết kế mở rộng được |
| PO — Color | 1 PO = 1 Color duy nhất trong 1 file PDF |
| Size matching | Auto-match sizes PDF↔Excel + thông báo size lệch |
| Nút trên UI | Vị trí đầu tiên, style xanh lá nổi bật, ẩn nút "Quét Sizes" |
| Preview dialog | Cho phép sửa PO, Color, Size (format lệch). Checkbox chọn/bỏ size |
| Ghi Excel | User review xong → ghi PO (cột A), Color (cột E), tick sizes + fill qty trên main UI |
| Progress | Progress bar theo % với checklist từng bước, lỗi dừng tại bước đó + nút Thử lại |

## Kiến trúc

```
excel_automation/
  pdf_po_parser.py          <- MỚI: Parse PDF, trích xuất dữ liệu

ui/
  pdf_import_dialog.py      <- MỚI: Preview dialog + Progress dialog
  excel_realtime_controller.py  <- SỬA: Thêm nút, ẩn Quét Sizes, thêm method _import_po_from_pdf

data/template_configs/
  pdf_import_config.json    <- MỚI: Config lưu vị trí trường trong PDF (mở rộng format sau)
```

## Module 1: PDFPOParser (excel_automation/pdf_po_parser.py)

### Dataclass output

```python
@dataclass
class PDFPOData:
    raw_po: str                      # "0009013330-1"
    po_number: str                   # "9013330"
    color_code: str                  # "3104"
    size_quantities: Dict[str, int]  # {"046": 60, "048": 140, ...}
    total_quantity: int              # 1030
    source_file: str                 # "Test.pdf"
```

### Luồng parse

1. Mở PDF bằng `pdfplumber`, đọc text tất cả trang
2. Trích PO: Regex tìm pattern `P. O. No.` hoặc tương đương → lấy giá trị (vd: `0009013330-1`) → bỏ leading zeros + bỏ phần sau dấu `-` → `9013330`
3. Trích Color: Từ dòng Article No. đầu tiên (vd: `62183104046`) → ký tự vị trí 5→8 (0-indexed: 4→8) → `3104`
4. Trích Size + Qty: Mỗi dòng item có pattern `Size:XX` kèm `Qty` → normalize size thành 3 chữ số (46→`046`, 96→`096`, 100 giữ nguyên `100`) → dict mapping {size: qty}
5. Tính total_quantity = sum(qty)
6. Trả về `PDFPOData`

### Quy tắc normalize size

- Size < 100: zero-pad 3 chữ số (46 → `046`, 96 → `096`)
- Size >= 100: giữ nguyên (`100`, `104`, `148`)
- Letter sizes (nếu có): giữ nguyên (`XS`, `M`, `XL`)

### Error handling

- Không tìm được PO → `RuntimeError("Không tìm thấy PO Number trong file PDF")`
- Không tìm được Article No. → `RuntimeError("Không tìm thấy Article Number trong file PDF")`
- Không tìm được dòng size nào → `RuntimeError("Không tìm thấy dữ liệu Size/Quantity trong file PDF")`

## Module 2: PDFImportDialog (ui/pdf_import_dialog.py)

### Input

- `parent: tk.Tk`
- `pdf_data: PDFPOData`
- `available_sizes: List[str]` (sizes đang có trong Excel)
- `on_confirm_callback: Callable[[str, str, Dict[str, int]], None]` — nhận (po, color, {size: qty})

### Layout

```
┌─────────────────────────────────────────────┐
│         📄 Import PO từ PDF                 │
│         File: Test.pdf                      │
├─────────────────────────────────────────────┤
│  PO Number    [input: 9013330 — editable]   │
│  Color Code   [input: 3104 — editable]      │
│  Total Qty:   1,030 (readonly)              │
├─────────────────────────────────────────────┤
│  📋 Chi tiết Size — Quantity:               │
│  ┌─────────────────────────────────────┐    │
│  │ ☑ │ Size [input] │ Qty  │ Trạng thái│   │
│  │ ☑ │ [046]  ro    │  60  │ ✅ Khớp   │   │
│  │ ☑ │ [048]  ro    │ 140  │ ✅ Khớp   │   │
│  │ ☐ │ [096]  edit  │  20  │ ⚠️ Chỉ PDF│   │
│  └─────────────────────────────────────┘    │
├─────────────────────────────────────────────┤
│  ⚠️ 2 size chỉ có trong PDF: 096, 100      │
│  ℹ️ 3 size chỉ có trong Excel: 038, 040    │
├─────────────────────────────────────────────┤
│              [Hủy]  [✅ Xác nhận & Ghi]    │
└─────────────────────────────────────────────┘
```

### Hành vi tương tác

| Thành phần | Hành vi |
|---|---|
| PO input | Pre-fill từ parser, user có thể sửa |
| Color input | Pre-fill từ parser, user có thể sửa |
| Checkbox | Size khớp Excel → auto checked. Size chỉ có PDF → unchecked mặc định |
| Size input (khớp) | Readonly — không cần sửa |
| Size input (không khớp) | Editable — user sửa format (vd: `96` → `096`) |
| Trạng thái realtime | Khi user sửa size input → kiểm tra lại có khớp Excel → cập nhật ✅/⚠️ + auto check nếu khớp |
| Nút Xác nhận | Collect PO, Color, các size đang checked + qty → gọi on_confirm_callback |
| Nút Hủy | Đóng dialog, không làm gì |

## Module 3: Progress Dialog (ui/pdf_import_dialog.py — cùng file)

### Khi nào hiển thị

Sau khi user bấm "✅ Xác nhận & Ghi" trên Preview Dialog.

### Các bước + trọng số %

| Bước | Mô tả hiển thị | % start → end |
|---|---|---|
| 1 | Đang đọc file PDF... | 0% → 20% |
| 2 | Đang trích xuất dữ liệu PO, Color, Sizes... | 20% → 35% |
| 3 | Đang scan sizes từ Excel... | 35% → 50% |
| 4 | Đang ghi PO vào Excel... | 50% → 65% |
| 5 | Đang ghi Color Code vào Excel... | 65% → 80% |
| 6 | Đang cập nhật Sizes & Quantities... | 80% → 95% |
| 7 | Hoàn tất! | 95% → 100% |

### Layout

```
┌──────────────────────────────────────┐
│       📄 Đang Import PO từ PDF       │
│                                      │
│  ████████████████░░░░░░░░  65%       │
│                                      │
│  ✅ Đọc file PDF                     │
│  ✅ Trích xuất dữ liệu              │
│  ✅ Scan sizes từ Excel              │
│  🔄 Đang ghi PO vào Excel...        │
│  ⬚ Ghi Color Code                   │
│  ⬚ Cập nhật Sizes & Quantities      │
│  ⬚ Hoàn tất                         │
└──────────────────────────────────────┘
```

### Hành vi

- Modal dialog — chặn thao tác app khi đang xử lý
- Mỗi bước hoàn thành → ✅ tick xanh, progress bar nhảy %
- Bước đang chạy → 🔄 + text "Đang..."
- Bước chưa chạy → ⬚ xám
- Hoàn tất 100% → tự đóng sau 1 giây + messagebox thành công

### Xử lý lỗi

- Bước nào fail → dừng tại bước đó, hiển thị ❌ + message lỗi chi tiết
- Nút "🔄 Thử lại" xuất hiện — chạy lại từ bước bị lỗi (không chạy lại bước đã thành công)
- Nút "Đóng" để user thoát nếu không muốn thử lại
- Các bước đã ghi thành công trước đó giữ nguyên (không rollback)

## Module 4: Tích hợp Main UI (ui/excel_realtime_controller.py)

### Thay đổi buttons_config

```python
buttons_config = [
    ("📄 Import PO từ PDF", self._import_po_from_pdf),  # MỚI — đầu tiên
    # ("🔍 Quét Sizes", self._scan_sizes),              # ẨN — không còn cần
    ("👁️ Ẩn dòng ngay", self._hide_rows_realtime),
    ("👁️‍🗨️ Hiện tất cả", self._show_all_rows),
    ("📝 Nhập Số Lượng Size", self._input_size_quantities),
    ("💾 Ghi vào Excel", self._write_quantities_to_excel),
    ("📦 Xuất Danh Sách Thùng", self._export_box_list),
    ("📄 Đọc PDF", self._open_pdf_reader),
]
```

Nút "Import PO từ PDF" dùng style riêng màu xanh lá (`Green.TButton`), tách biệt khỏi buttons_config loop, tương tự cách Update PO/Color dùng style vàng.

### Luồng _import_po_from_pdf()

```
User bấm nút
    ↓
[1] Kiểm tra đã mở Excel chưa → nếu chưa: warning return
    ↓
[2] filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    ↓
[3] Hiển thị Progress Dialog bước 1-2: Parse PDF → PDFPOData
    ↓
[4] Progress bước 3: Auto scan sizes từ Excel (com_manager.scan_sizes())
    ↓
[5] Tạm dừng progress → Mở PDFImportDialog(pdf_data, available_sizes, on_confirm)
    ↓
[6] User review, sửa, bấm "Xác nhận"
    ↓
[7] Progress tiếp tục bước 4-7: Ghi PO → Color → Sizes/Qty → Hoàn tất
    ↓
[8] Cập nhật main UI: tick checkboxes, fill quantity entries, update PO/Color display
```

### Callback on_confirm ghi Excel

| Bước | Hành động | Module sử dụng |
|---|---|---|
| Ghi PO | Cột A, dòng 19→end | `POUpdateManager.update_po_bulk()` |
| Ghi Color | Cột E, dòng 19→end (prefix `'`) | `ColorCodeUpdateManager.update_color_code_bulk()` |
| Tick sizes | Set `self.checkboxes[size].set(True)` cho sizes đã checked | Direct UI update |
| Fill qty | `self.quantity_entries[size].insert(0, qty)` cho từng size | Direct UI update |
| Refresh display | `self._update_po_color_display()` + `self._update_box_count_display()` | Existing methods |

## Testing

### Unit test cho PDFPOParser

File: `tests/test_pdf_po_parser.py`

- Test parse file Test.pdf → kiểm tra PO, Color, sizes, quantities đúng
- Test PO extraction: `0009013330-1` → `9013330`
- Test Color extraction: `62183104046` → `3104`
- Test size normalization: 46→`046`, 96→`096`, 100→`100`
- Test error handling: file không tồn tại, PDF không có PO, PDF rỗng

### Manual test cho UI

- Import PDF → preview hiển thị đúng
- Sửa PO/Color trong preview → giá trị mới được ghi
- Sửa size không khớp → trạng thái cập nhật realtime
- Bấm Xác nhận → progress chạy đúng %, ghi đúng vào Excel
- Test lỗi giữa chừng → dừng đúng bước, thử lại hoạt động

## Phụ thuộc

- `pdfplumber` (đã có trong requirements.txt)
- Các module hiện có: `POUpdateManager`, `ColorCodeUpdateManager`, `ExcelCOMManager`
- Không cần thêm thư viện mới
