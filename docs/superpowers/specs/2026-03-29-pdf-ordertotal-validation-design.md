# PDF Ordertotal Validation

## Vấn đề

Khi import PDF, `total_quantity` trong `PDFPOData` được tính bằng `sum(size_quantities.values())` — tức chỉ cộng các size qty mà regex parse được. Nếu regex bỏ sót một số dòng size (do format khác biệt), tổng vẫn "đúng" theo dữ liệu đã parse, nhưng thực tế thiếu qty so với PDF gốc.

PDF Purchase Order có dòng `Ordertotal` riêng chứa tổng qty chính thức. Cần đọc giá trị này và cross-check với tổng parse được.

## Giải pháp

Validation trong Parser + Warning trong Dialog.

### Thay đổi PDFPOData

Thêm 2 field:
- `ordertotal_from_pdf: Optional[int] = None` — giá trị Ordertotal gốc đọc từ PDF. None nếu không tìm thấy dòng Ordertotal.
- `quantity_mismatch: bool = False` — True khi `sum(size_quantities) != ordertotal_from_pdf`.

### Thay đổi PDFPOParser

Thêm method `_extract_ordertotal(full_text: str) -> Optional[int]`:
- Regex pattern: `r'Ordertotal\s+(\d+)\s+'`
- Match dòng dạng: `Ordertotal 1030 20898.70 USD`
- Trả `None` nếu không match (không raise error — đây là validation phụ).

Trong `parse()`:
- Gọi `_extract_ordertotal()` sau khi parse xong size_quantities.
- `parsed_total = sum(size_quantities.values())`
- `mismatch = ordertotal is not None and parsed_total != ordertotal`
- `total_quantity` vẫn dùng `parsed_total` (không dùng ordertotal).
- Log warning nếu mismatch.

### Thay đổi PDFImportDialog

Khi `pdf_data.quantity_mismatch == True`, hiển thị warning frame trong info section:
- Nằm ngay dưới dòng Total Qty hiện có.
- Màu đỏ/cam, format: `⚠️ Chênh lệch! Parse được {parsed_total} qty từ {num_sizes} size, Ordertotal PDF = {ordertotal_from_pdf} (thiếu {diff} qty)`
- Khi `ordertotal_from_pdf is None`: hiển thị info nhẹ `ℹ️ Không tìm thấy dòng Ordertotal trong PDF để kiểm tra chéo`
- Không block import — user tự quyết định.

### File thay đổi

| Action | File |
|--------|------|
| Sửa | `excel_automation/pdf_po_parser.py` |
| Sửa | `ui/pdf_import_dialog.py` |
| Sửa | `tests/test_pdf_po_parser.py` |

### Test cases

**Parser tests:**
- `_extract_ordertotal` với text có `Ordertotal 1030 20898.70 USD` → 1030
- `_extract_ordertotal` với text không có Ordertotal → None
- `parse()` trả `quantity_mismatch=False` khi tổng khớp
- `parse()` trả `quantity_mismatch=True` khi tổng không khớp
- `parse()` trả `quantity_mismatch=False` khi không tìm thấy Ordertotal (None)
- Full parse Test.pdf → `ordertotal_from_pdf=1030`, `quantity_mismatch=False`

**Dialog:** Không cần test tự động (tkinter), kiểm tra thủ công.
