# Plan: PDF Ordertotal Validation

**Spec:** `docs/superpowers/specs/2026-03-29-pdf-ordertotal-validation-design.md`

**Tech Stack:** Python 3, pdfplumber, tkinter, pytest

---

## File Structure

| Action | File | Responsibility |
|--------|------|----------------|
| Sửa | `excel_automation/pdf_po_parser.py` | Thêm `_extract_ordertotal`, thêm field vào PDFPOData, cập nhật `parse()` |
| Sửa | `tests/test_pdf_po_parser.py` | Thêm test class cho ordertotal extraction + mismatch detection |
| Sửa | `ui/pdf_import_dialog.py` | Hiển thị warning khi mismatch hoặc info khi không tìm thấy Ordertotal |

---

### Task 1: Parser — Thêm `_extract_ordertotal` và cập nhật PDFPOData + parse()

**File:** `excel_automation/pdf_po_parser.py`

**Bước 1:** Thêm 2 field vào `PDFPOData`:
```python
ordertotal_from_pdf: Optional[int] = None
quantity_mismatch: bool = False
```
Import thêm `Optional` từ `typing`.

**Bước 2:** Thêm static method `_extract_ordertotal`:
```python
@staticmethod
def _extract_ordertotal(full_text: str) -> Optional[int]:
    pattern = r'Ordertotal\s+(\d+)\s+'
    match = re.search(pattern, full_text)
    if not match:
        logger.warning("Không tìm thấy dòng Ordertotal trong PDF")
        return None
    return int(match.group(1))
```

**Bước 3:** Cập nhật `parse()` — sau dòng `total_quantity = sum(...)`:
```python
ordertotal = PDFPOParser._extract_ordertotal(full_text)
mismatch = ordertotal is not None and total_quantity != ordertotal
if mismatch:
    logger.warning(
        f"Chênh lệch qty! Parse={total_quantity}, Ordertotal PDF={ordertotal}, "
        f"thiếu={ordertotal - total_quantity}"
    )
```
Set `ordertotal_from_pdf=ordertotal` và `quantity_mismatch=mismatch` trong return.

**Verification:** `pytest tests/test_pdf_po_parser.py -v` — tests cũ phải pass, test mới pass.

---

### Task 2: Tests cho ordertotal validation

**File:** `tests/test_pdf_po_parser.py`

Thêm class `TestOrdertotalExtraction`:
- `test_extract_ordertotal_standard` — text có `Ordertotal 1030 20898.70 USD` → 1030
- `test_extract_ordertotal_not_found` — text không có Ordertotal → None
- `test_extract_ordertotal_different_number` — `Ordertotal 500 10000.00 USD` → 500

Thêm class `TestQuantityMismatch`:
- `test_no_mismatch_when_totals_match` — mock parse scenario tổng khớp → `quantity_mismatch=False`
- `test_mismatch_when_totals_differ` — mock parse scenario tổng khác → `quantity_mismatch=True`
- `test_no_mismatch_when_ordertotal_not_found` — không có Ordertotal → `quantity_mismatch=False`, `ordertotal_from_pdf=None`

Cập nhật `TestFullParse.test_parse_test_pdf`:
- Assert thêm `result.ordertotal_from_pdf == 1030`
- Assert thêm `result.quantity_mismatch == False`

**Verification:** `pytest tests/test_pdf_po_parser.py -v`

---

### Task 3: Dialog — Hiển thị warning/info

**File:** `ui/pdf_import_dialog.py`

Trong `_create_widgets()`, sau dòng hiển thị Total Qty, thêm logic:
- Nếu `self.pdf_data.quantity_mismatch == True`:
  - Thêm frame warning màu đỏ
  - Text: `⚠️ Chênh lệch! Parse được {total_quantity} qty từ {len(size_quantities)} size, Ordertotal PDF = {ordertotal_from_pdf} (thiếu {diff} qty)`
- Nếu `self.pdf_data.ordertotal_from_pdf is None`:
  - Thêm label info màu xám
  - Text: `ℹ️ Không tìm thấy dòng Ordertotal trong PDF để kiểm tra chéo`
- Nếu match (không mismatch, có ordertotal): không hiển thị gì thêm (hoặc ✅ nhỏ).

**Verification:** Chạy `python excel_realtime_controller.py`, import Test.pdf, kiểm tra dialog hiển thị đúng.

---

## Task Dependencies

```
Task 1 (Parser) → Task 2 (Tests) → Task 3 (Dialog)
```

Task 1 và 2 có thể chạy song song (viết test trước rồi implement đều được), nhưng vì scope nhỏ nên chạy tuần tự cho đơn giản. Task 3 phụ thuộc Task 1 (cần field mới trong PDFPOData).
