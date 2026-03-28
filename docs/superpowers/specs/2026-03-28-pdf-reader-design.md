# PDF Reader Feature — Design Spec

**Date:** 2026-03-28
**Approach:** pdfplumber (digital PDF) + Tesseract OCR (scanned PDF)
**Status:** Approved

---

## 1. Mục tiêu

Thêm tính năng đọc PDF vào ứng dụng Excel Real-Time Controller. Giai đoạn 1 (proof of concept): extract text từ file PDF (cả digital lẫn scan) và hiển thị kết quả trên giao diện để đánh giá chất lượng.

## 2. Kiến trúc Module

```
PythonExcel/
├── excel_automation/
│   └── pdf_reader.py          # Core logic — extract text từ PDF
├── ui/
│   └── pdf_reader_dialog.py   # Dialog Tkinter — chọn file, hiển thị text
```

### pdf_reader.py — Module logic thuần (không phụ thuộc UI)

- `extract_text_from_pdf(file_path: str, on_progress: Callable = None) -> str` — hàm chính, trả về toàn bộ text từ PDF. `on_progress(page_num, total_pages, is_ocr)` callback để UI cập nhật trạng thái
- `is_scanned_page(page) -> bool` — kiểm tra trang có text hay chỉ là ảnh
- `extract_page_text(page, page_number: int) -> str` — extract text 1 trang (digital hoặc OCR)
- `check_ocr_available() -> dict` — kiểm tra Tesseract + Poppler đã cài chưa, trả về `{"tesseract": bool, "poppler": bool}`
- Tự detect từng trang: digital → pdfplumber, scanned → pdf2image + pytesseract

### pdf_reader_dialog.py — Popup dialog

- Nút "Chọn file PDF" mở `filedialog.askopenfilename()`
- Hiển thị text kết quả trong `ScrolledText` widget (readonly, có scrollbar)
- Hiển thị progress/trạng thái khi đang xử lý
- Nút được thêm vào giao diện chính `ui/excel_realtime_controller.py`

## 3. Logic xử lý PDF

```
PDF file
  │
  ▼
Mở file bằng pdfplumber
  │
  ▼
Duyệt từng trang (page)
  │
  ├─ page.extract_text() có text (>= 10 ký tự)? ── YES ──▶ Lấy text trực tiếp
  │
  └─ Không có text (scanned) ──▶ Convert trang → ảnh (pdf2image)
                                       │
                                       ▼
                                 pytesseract OCR (lang='vie')
                                       │
                                       ▼
                                   Trả về text
  │
  ▼
Gộp text tất cả trang, đánh dấu "--- Trang X ---"
  │
  ▼
Trả kết quả về cho UI
```

### Detect trang scan

- Gọi `page.extract_text()` trước
- Nếu trả về chuỗi rỗng, toàn khoảng trắng, hoặc ít hơn 10 ký tự → coi là trang scan
- Ngưỡng 10 ký tự để tránh trường hợp có vài ký tự rác

### Xử lý lỗi

| Trường hợp | Xử lý |
|---|---|
| File không tồn tại | Raise `FileNotFoundError` |
| File không phải PDF | Raise `ValueError` |
| Tesseract chưa cài | Trả text `"[Trang X: Cần cài Tesseract OCR để đọc trang scan]"` — không crash |
| Poppler chưa cài | Tương tự — thông báo cần cài, không crash |
| PDF bị password | Bắt exception, thông báo rõ cho user |
| Trang digital vẫn đọc được bình thường khi thiếu Tesseract/Poppler |

## 4. Giao diện Dialog

```
┌─────────────────────────────────────────────────────┐
│  Đọc PDF                                    [X]    │
├─────────────────────────────────────────────────────┤
│                                                     │
│  [Chọn file PDF]    D:\path\to\file.pdf             │
│                                                     │
│  ┌─────────────────────────────────────────────┐    │
│  │ --- Trang 1 ---                              │    │
│  │ Nội dung text extract được từ trang 1...     │    │
│  │                                              │    │
│  │ --- Trang 2 ---                              │    │
│  │ Nội dung text extract được từ trang 2...     │    │
│  │                                         ▼    │    │
│  └─────────────────────────────────────────────┘    │
│                                                     │
│  Trạng thái: Đã đọc 5/5 trang (2 trang OCR)       │
│                                                     │
│  [Copy text]                       [Đóng]           │
│                                                     │
└─────────────────────────────────────────────────────┘
```

### Widgets

| Widget | Mô tả |
|---|---|
| Nút "Chọn file PDF" | Mở filedialog, filter `*.pdf` |
| Label path | Hiện đường dẫn file đã chọn |
| ScrolledText | Vùng hiển thị text, readonly, scrollbar, font monospace |
| Label trạng thái | Tiến trình: "Đang đọc trang 3/10...", "Đang OCR trang 5..." |
| Nút "Copy text" | Copy toàn bộ text vào clipboard |
| Nút "Đóng" | Đóng dialog |

### UX khi OCR chậm

- Chạy extract trên **thread riêng** (`threading.Thread`) để UI không bị đơ
- Cập nhật trạng thái realtime qua `root.after()` — pattern đã dùng trong project
- Nút "Chọn file PDF" bị disable trong lúc đang xử lý

### Kích thước dialog

- 700x500 pixels, resizable
- Nhớ vị trí qua `dialog_config_manager.py` (pattern có sẵn trong project)

## 5. Dependencies

### Thư viện Python (thêm vào requirements.txt)

```
pdfplumber>=0.11.0     # Extract text từ digital PDF
pdf2image>=1.17.0      # Convert PDF page → ảnh cho OCR
pytesseract>=0.3.10    # Python wrapper cho Tesseract OCR
Pillow>=10.0.0         # Đã có sẵn (dependency của pdfplumber)
```

### Phần mềm bên ngoài (Windows)

| Software | Cài đặt | Ghi chú |
|---|---|---|
| Tesseract OCR | Download `.exe` từ UB Mannheim | Chọn thêm language pack Vietnamese khi cài |
| Poppler | Download binary, giải nén, thêm `bin/` vào PATH | Cần cho pdf2image convert PDF → ảnh |

### Graceful degradation

- App khởi động bình thường — tính năng PDF không ảnh hưởng các tính năng Excel hiện có
- PDF digital (có text sẵn) đọc được mà không cần Tesseract/Poppler
- Nếu thiếu Tesseract/Poppler → hiện thông báo hướng dẫn cài đặt, không crash

### Đóng gói PyInstaller

- Thêm `pdfplumber`, `pdf2image`, `pytesseract` vào hidden imports trong `.spec`
- Tesseract + Poppler cần bundle kèm hoặc hướng dẫn user cài riêng
- Dự kiến tăng kích thước `.exe` thêm ~15-20MB (không tính Tesseract/Poppler)

## 6. Tích hợp vào app chính

- Thêm nút "Đọc PDF" vào giao diện chính `ui/excel_realtime_controller.py`
- Import `PdfReaderDialog` từ `ui/pdf_reader_dialog.py`
- Nút mở dialog dạng `Toplevel` (không block cửa sổ chính)
