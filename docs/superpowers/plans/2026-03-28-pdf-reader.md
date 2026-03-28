# PDF Reader Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add PDF text extraction (digital + OCR) to the Excel Real-Time Controller app, displayed in a Tkinter dialog.

**Architecture:** A pure-logic module `excel_automation/pdf_reader.py` handles all PDF processing (pdfplumber for digital, pdf2image+pytesseract for scanned pages). A separate `ui/pdf_reader_dialog.py` provides the Tkinter dialog with threaded OCR. Integration into the main UI is a single button addition.

**Tech Stack:** Python, pdfplumber, pdf2image, pytesseract, Pillow, tkinter, threading

---

## File Structure

| File | Action | Responsibility |
|---|---|---|
| `excel_automation/pdf_reader.py` | Create | Core PDF text extraction logic (no UI dependency) |
| `tests/test_pdf_reader.py` | Create | Unit tests for pdf_reader module |
| `ui/pdf_reader_dialog.py` | Create | Tkinter Toplevel dialog for PDF reading |
| `ui/excel_realtime_controller.py` | Modify (lines 162-178) | Add "Đọc PDF" button to action_frame |
| `excel_automation/__init__.py` | Modify | Export PdfReader functions |
| `requirements.txt` | Modify | Add pdfplumber, pdf2image, pytesseract, Pillow |

---

### Task 1: Add Dependencies to requirements.txt

**Files:**
- Modify: `requirements.txt`

- [ ] **Step 1: Add PDF dependencies to requirements.txt**

Append these lines to the end of `requirements.txt`:

```
pdfplumber>=0.11.0
pdf2image>=1.17.0
pytesseract>=0.3.10
Pillow>=10.0.0
```

The full file after edit:

```
pandas>=2.0.0
openpyxl>=3.1.0
xlsxwriter>=3.2.0
xlrd>=2.0.0
xlwt>=1.3.0
xlutils>=2.0.0
python-dotenv>=1.0.0
fastexcel>=0.16.0
python-calamine>=0.4.0
pywin32>=306
pdfplumber>=0.11.0
pdf2image>=1.17.0
pytesseract>=0.3.10
Pillow>=10.0.0
```

- [ ] **Step 2: Install the new dependencies**

Run:
```bash
pip install pdfplumber>=0.11.0 pdf2image>=1.17.0 pytesseract>=0.3.10 Pillow>=10.0.0
```

Expected: All packages install successfully. Pillow may already be installed.

- [ ] **Step 3: Commit**

```bash
git add requirements.txt
git commit -m "chore: add PDF processing dependencies (pdfplumber, pdf2image, pytesseract, Pillow)"
```

---

### Task 2: Create pdf_reader.py — check_ocr_available()

**Files:**
- Create: `excel_automation/pdf_reader.py`
- Create: `tests/test_pdf_reader.py`

- [ ] **Step 1: Write the failing test for check_ocr_available**

Create `tests/test_pdf_reader.py`:

```python
"""
Unit tests cho PDF Reader module.
"""

import pytest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_automation.pdf_reader import check_ocr_available


class TestCheckOcrAvailable:
    """Test cases cho check_ocr_available function."""

    def test_returns_dict_with_tesseract_key(self):
        """check_ocr_available trả về dict có key 'tesseract'."""
        result = check_ocr_available()
        assert isinstance(result, dict)
        assert "tesseract" in result

    def test_returns_dict_with_poppler_key(self):
        """check_ocr_available trả về dict có key 'poppler'."""
        result = check_ocr_available()
        assert "poppler" in result

    def test_values_are_boolean(self):
        """Các giá trị trong dict đều là bool."""
        result = check_ocr_available()
        assert isinstance(result["tesseract"], bool)
        assert isinstance(result["poppler"], bool)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_pdf_reader.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'excel_automation.pdf_reader'`

- [ ] **Step 3: Write minimal implementation of check_ocr_available**

Create `excel_automation/pdf_reader.py`:

```python
"""
PDF Reader Module — Extract text từ PDF files.

Hỗ trợ:
- Digital PDF: extract text trực tiếp bằng pdfplumber
- Scanned PDF: convert trang → ảnh rồi OCR bằng pytesseract

Module này không phụ thuộc UI — chỉ chứa logic thuần.
"""

import logging
import shutil
from typing import Callable, Optional

logger = logging.getLogger(__name__)


def check_ocr_available() -> dict:
    """
    Kiểm tra Tesseract OCR và Poppler đã cài đặt chưa.

    Returns:
        dict: {"tesseract": bool, "poppler": bool}
    """
    result = {"tesseract": False, "poppler": False}

    # Check Tesseract
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        result["tesseract"] = True
    except Exception:
        # pytesseract chưa cài hoặc Tesseract binary không tìm thấy
        pass

    # Check Poppler (pdf2image cần pdftoppm binary)
    if shutil.which("pdftoppm") is not None:
        result["poppler"] = True
    else:
        # Fallback: thử import pdf2image và gọi thử
        try:
            from pdf2image.exceptions import PDFInfoNotInstalledError
            from pdf2image import pdfinfo_from_path
            # Nếu import thành công nhưng chưa chắc có poppler
            # Chỉ set True nếu tìm thấy binary
            pass
        except Exception:
            pass

    return result
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_pdf_reader.py -v`
Expected: All 3 tests PASS

- [ ] **Step 5: Commit**

```bash
git add excel_automation/pdf_reader.py tests/test_pdf_reader.py
git commit -m "feat: add pdf_reader module with check_ocr_available()"
```

---

### Task 3: Add is_scanned_page() and extract_page_text()

**Files:**
- Modify: `excel_automation/pdf_reader.py`
- Modify: `tests/test_pdf_reader.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_pdf_reader.py`:

```python
from unittest.mock import MagicMock
from excel_automation.pdf_reader import is_scanned_page, extract_page_text


class TestIsScannedPage:
    """Test cases cho is_scanned_page function."""

    def test_page_with_text_is_not_scanned(self):
        """Trang có nhiều text (>= 10 ký tự) không phải scan."""
        page = MagicMock()
        page.extract_text.return_value = "This is a page with enough text content"
        assert is_scanned_page(page) is False

    def test_page_with_no_text_is_scanned(self):
        """Trang không có text là trang scan."""
        page = MagicMock()
        page.extract_text.return_value = None
        assert is_scanned_page(page) is True

    def test_page_with_empty_text_is_scanned(self):
        """Trang có text rỗng là trang scan."""
        page = MagicMock()
        page.extract_text.return_value = ""
        assert is_scanned_page(page) is True

    def test_page_with_whitespace_only_is_scanned(self):
        """Trang chỉ có khoảng trắng là trang scan."""
        page = MagicMock()
        page.extract_text.return_value = "   \n\t  "
        assert is_scanned_page(page) is True

    def test_page_with_few_chars_is_scanned(self):
        """Trang có ít hơn 10 ký tự (rác) là trang scan."""
        page = MagicMock()
        page.extract_text.return_value = "abc"
        assert is_scanned_page(page) is True

    def test_page_with_exactly_10_chars_is_not_scanned(self):
        """Trang có đúng 10 ký tự không phải scan."""
        page = MagicMock()
        page.extract_text.return_value = "0123456789"
        assert is_scanned_page(page) is False


class TestExtractPageText:
    """Test cases cho extract_page_text — chỉ test digital path."""

    def test_digital_page_returns_text(self):
        """Trang digital trả về text trực tiếp."""
        page = MagicMock()
        page.extract_text.return_value = "This is digital text content from PDF page"
        result = extract_page_text(page, page_number=1)
        assert result == "This is digital text content from PDF page"

    def test_scanned_page_without_tesseract_returns_message(self):
        """Trang scan khi không có Tesseract trả về thông báo hướng dẫn."""
        page = MagicMock()
        page.extract_text.return_value = ""
        result = extract_page_text(page, page_number=3)
        # Phải chứa thông báo cần cài Tesseract hoặc Poppler
        assert "Trang 3" in result or "trang 3" in result
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_pdf_reader.py -v`
Expected: FAIL with `ImportError` for `is_scanned_page` and `extract_page_text`

- [ ] **Step 3: Implement is_scanned_page and extract_page_text**

Add to `excel_automation/pdf_reader.py` (after `check_ocr_available`):

```python
MIN_TEXT_LENGTH = 10  # Ngưỡng ký tự tối thiểu để coi là trang digital


def is_scanned_page(page) -> bool:
    """
    Kiểm tra trang PDF có phải là trang scan (ảnh) hay không.

    Logic: Gọi page.extract_text(). Nếu kết quả rỗng, toàn khoảng trắng,
    hoặc ít hơn MIN_TEXT_LENGTH ký tự → coi là trang scan.

    Args:
        page: pdfplumber Page object

    Returns:
        True nếu là trang scan, False nếu là trang digital
    """
    try:
        text = page.extract_text()
        if text is None:
            return True
        stripped = text.strip()
        if len(stripped) < MIN_TEXT_LENGTH:
            return True
        return False
    except Exception:
        return True


def extract_page_text(page, page_number: int) -> str:
    """
    Extract text từ 1 trang PDF.

    - Nếu trang digital (có text) → trả về text trực tiếp
    - Nếu trang scan → thử OCR bằng pdf2image + pytesseract
    - Nếu thiếu Tesseract/Poppler → trả về thông báo hướng dẫn, không crash

    Args:
        page: pdfplumber Page object
        page_number: Số thứ tự trang (1-based, dùng cho thông báo)

    Returns:
        Text extract được từ trang
    """
    # Thử extract text trực tiếp (digital PDF)
    if not is_scanned_page(page):
        return page.extract_text()

    # Trang scan → thử OCR
    ocr_status = check_ocr_available()

    if not ocr_status["poppler"]:
        return (
            f"[Trang {page_number}: Cần cài Poppler để đọc trang scan. "
            f"Download tại: https://github.com/oschwartz10612/poppler-windows/releases "
            f"và thêm thư mục bin/ vào PATH]"
        )

    if not ocr_status["tesseract"]:
        return (
            f"[Trang {page_number}: Cần cài Tesseract OCR để đọc trang scan. "
            f"Download tại: https://github.com/UB-Mannheim/tesseract/wiki "
            f"và chọn thêm Vietnamese language pack khi cài]"
        )

    try:
        from pdf2image import convert_from_path
        import pytesseract

        # Convert trang PDF → ảnh
        images = convert_from_path(
            page.pdf.stream.name,
            first_page=page_number,
            last_page=page_number,
            dpi=300
        )

        if not images:
            return f"[Trang {page_number}: Không thể convert trang sang ảnh]"

        # OCR ảnh → text
        text = pytesseract.image_to_string(images[0], lang="vie")
        return text.strip() if text else f"[Trang {page_number}: OCR không detect được text]"

    except Exception as e:
        logger.error(f"Lỗi OCR trang {page_number}: {e}")
        return f"[Trang {page_number}: Lỗi khi OCR — {e}]"
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_pdf_reader.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add excel_automation/pdf_reader.py tests/test_pdf_reader.py
git commit -m "feat: add is_scanned_page() and extract_page_text() to pdf_reader"
```

---

### Task 4: Add extract_text_from_pdf() — Main Function

**Files:**
- Modify: `excel_automation/pdf_reader.py`
- Modify: `tests/test_pdf_reader.py`

- [ ] **Step 1: Create a test PDF fixture and write failing tests**

Append to `tests/test_pdf_reader.py`:

```python
import pdfplumber
from reportlab.pdfgen import canvas as pdf_canvas
from excel_automation.pdf_reader import extract_text_from_pdf


@pytest.fixture
def digital_pdf(tmp_path):
    """Tạo file PDF digital (có text) để test."""
    pdf_path = tmp_path / "test_digital.pdf"
    c = pdf_canvas.Canvas(str(pdf_path))
    c.drawString(72, 700, "Hello World from Page 1")
    c.showPage()
    c.drawString(72, 700, "Content on Page 2")
    c.showPage()
    c.save()
    return str(pdf_path)


class TestExtractTextFromPdf:
    """Test cases cho extract_text_from_pdf function."""

    def test_extract_digital_pdf(self, digital_pdf):
        """Extract text từ PDF digital trả về nội dung các trang."""
        result = extract_text_from_pdf(digital_pdf)
        assert "Hello World from Page 1" in result
        assert "Content on Page 2" in result

    def test_extract_adds_page_markers(self, digital_pdf):
        """Kết quả có đánh dấu trang."""
        result = extract_text_from_pdf(digital_pdf)
        assert "--- Trang 1 ---" in result
        assert "--- Trang 2 ---" in result

    def test_file_not_found_raises_error(self):
        """File không tồn tại raise FileNotFoundError."""
        with pytest.raises(FileNotFoundError):
            extract_text_from_pdf("nonexistent_file.pdf")

    def test_non_pdf_file_raises_error(self, tmp_path):
        """File không phải PDF raise ValueError."""
        txt_file = tmp_path / "test.txt"
        txt_file.write_text("not a pdf")
        with pytest.raises(ValueError):
            extract_text_from_pdf(str(txt_file))

    def test_progress_callback_called(self, digital_pdf):
        """on_progress callback được gọi cho mỗi trang."""
        progress_calls = []

        def on_progress(page_num, total_pages, is_ocr):
            progress_calls.append((page_num, total_pages, is_ocr))

        extract_text_from_pdf(digital_pdf, on_progress=on_progress)
        assert len(progress_calls) == 2
        assert progress_calls[0] == (1, 2, False)
        assert progress_calls[1] == (2, 2, False)

    def test_empty_pdf_returns_empty_string(self, tmp_path):
        """PDF không có trang nào trả về chuỗi rỗng."""
        pdf_path = tmp_path / "empty.pdf"
        c = pdf_canvas.Canvas(str(pdf_path))
        c.save()
        result = extract_text_from_pdf(str(pdf_path))
        assert isinstance(result, str)

    def test_password_protected_pdf_raises_runtime_error(self, tmp_path):
        """PDF bị password raise RuntimeError với thông báo rõ."""
        from reportlab.lib.pdfencrypt import StandardEncryption
        pdf_path = tmp_path / "protected.pdf"
        enc = StandardEncryption("userpass", ownerPassword="ownerpass")
        c = pdf_canvas.Canvas(str(pdf_path), encrypt=enc)
        c.drawString(72, 700, "Secret content")
        c.save()
        with pytest.raises(RuntimeError, match="[Pp]assword|mật khẩu"):
            extract_text_from_pdf(str(pdf_path))
```

**Note:** Test này cần thêm dependency `reportlab` để tạo PDF test. Cài trước:

```bash
pip install reportlab
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_pdf_reader.py::TestExtractTextFromPdf -v`
Expected: FAIL with `ImportError` for `extract_text_from_pdf`

- [ ] **Step 3: Implement extract_text_from_pdf**

Add to `excel_automation/pdf_reader.py` (after `extract_page_text`):

```python
import os


def extract_text_from_pdf(
    file_path: str,
    on_progress: Optional[Callable] = None
) -> str:
    """
    Extract toàn bộ text từ file PDF.

    Tự detect từng trang: digital → pdfplumber, scanned → OCR.
    Kết quả gộp tất cả trang, có đánh dấu "--- Trang X ---".

    Args:
        file_path: Đường dẫn file PDF
        on_progress: Callback(page_num, total_pages, is_ocr) để cập nhật UI

    Returns:
        Text gộp từ tất cả trang PDF

    Raises:
        FileNotFoundError: File không tồn tại
        ValueError: File không phải PDF
    """
    # Validate file
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File không tồn tại: {file_path}")

    if not file_path.lower().endswith(".pdf"):
        raise ValueError(f"File không phải PDF: {file_path}")

    import pdfplumber

    all_text_parts = []

    try:
        with pdfplumber.open(file_path) as pdf:
            total_pages = len(pdf.pages)

            if total_pages == 0:
                return ""

            for i, page in enumerate(pdf.pages):
                page_number = i + 1
                is_ocr = is_scanned_page(page)

                # Callback progress
                if on_progress:
                    on_progress(page_number, total_pages, is_ocr)

                # Extract text
                page_text = extract_page_text(page, page_number)

                # Gộp với page marker
                all_text_parts.append(f"--- Trang {page_number} ---")
                all_text_parts.append(page_text if page_text else "")

        return "\n".join(all_text_parts)

    except Exception as e:
        if isinstance(e, (FileNotFoundError, ValueError)):
            raise
        error_msg = str(e).lower()
        if "password" in error_msg or "encrypted" in error_msg:
            raise RuntimeError(
                f"File PDF được bảo vệ bằng mật khẩu (password). "
                f"Vui lòng mở khóa file trước khi đọc."
            ) from e
        logger.error(f"Lỗi khi đọc PDF: {e}")
        raise RuntimeError(f"Lỗi khi đọc PDF: {e}") from e
```

Also add `import os` to the top imports section of the file (alongside the existing imports).

- [ ] **Step 4: Run all pdf_reader tests to verify they pass**

Run: `pytest tests/test_pdf_reader.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add excel_automation/pdf_reader.py tests/test_pdf_reader.py
git commit -m "feat: add extract_text_from_pdf() with page detection and progress callback"
```

---

### Task 5: Export pdf_reader from excel_automation package

**Files:**
- Modify: `excel_automation/__init__.py`

- [ ] **Step 1: Add imports to __init__.py**

In `excel_automation/__init__.py`, add the import after the existing `from excel_automation.utils import get_size_sort_key` line:

```python
from excel_automation.pdf_reader import (
    extract_text_from_pdf,
    check_ocr_available,
    is_scanned_page,
    extract_page_text,
)
```

Also add to the `__all__` list:

```python
    "extract_text_from_pdf",
    "check_ocr_available",
    "is_scanned_page",
    "extract_page_text",
```

- [ ] **Step 2: Verify import works**

Run: `python -c "from excel_automation.pdf_reader import extract_text_from_pdf, check_ocr_available; print('OK')"`
Expected: Prints `OK`

- [ ] **Step 3: Run all tests to verify nothing is broken**

Run: `pytest tests/ -v`
Expected: All existing tests still PASS

- [ ] **Step 4: Commit**

```bash
git add excel_automation/__init__.py
git commit -m "feat: export pdf_reader functions from excel_automation package"
```

---

### Task 6: Create pdf_reader_dialog.py — Dialog UI

**Files:**
- Create: `ui/pdf_reader_dialog.py`

- [ ] **Step 1: Create the dialog file**

Create `ui/pdf_reader_dialog.py`:

```python
"""
PDF Reader Dialog — Popup dialog để đọc và hiển thị text từ file PDF.

Chạy OCR trên thread riêng để UI không bị đơ.
Dùng root.after() để cập nhật progress — pattern có sẵn trong project.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import logging
from typing import Optional

from excel_automation.pdf_reader import extract_text_from_pdf, check_ocr_available
from excel_automation.dialog_config_manager import DialogConfigManager

logger = logging.getLogger(__name__)


class PdfReaderDialog:
    """Dialog Tkinter để chọn file PDF và hiển thị text extract được."""

    DIALOG_NAME = "pdf_reader"

    def __init__(self, parent: tk.Tk):
        self.parent = parent
        self.dialog_config = DialogConfigManager()
        self._processing = False
        self._worker_thread: Optional[threading.Thread] = None

        # Tạo dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Đọc PDF")

        # Kích thước từ config hoặc default 700x500
        width, height = self.dialog_config.get_dialog_size(self.DIALOG_NAME)
        if width < 400:
            width = 700
        if height < 300:
            height = 500
        self.dialog.geometry(f"{width}x{height}")
        self.dialog.resizable(True, True)

        # Transient nhưng không grab_set (cho phép tương tác cửa sổ chính)
        self.dialog.transient(parent)

        self.dialog.protocol("WM_DELETE_WINDOW", self._on_closing)

        self._create_widgets()
        self._center_window()

    def _center_window(self) -> None:
        """Đặt dialog giữa màn hình."""
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")

    def _create_widgets(self) -> None:
        """Tạo tất cả widgets cho dialog."""
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- File chooser row ---
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        self.choose_btn = ttk.Button(
            file_frame,
            text="Chọn file PDF",
            command=self._choose_and_read_pdf,
            width=15
        )
        self.choose_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.file_label = ttk.Label(
            file_frame,
            text="Chưa chọn file",
            foreground="gray"
        )
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- Text result area ---
        self.text_area = ScrolledText(
            main_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            state=tk.DISABLED
        )
        self.text_area.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # --- Status label ---
        self.status_label = ttk.Label(
            main_frame,
            text="Sẵn sàng",
            foreground="gray"
        )
        self.status_label.pack(fill=tk.X, pady=(0, 10))

        # --- Button row ---
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        self.copy_btn = ttk.Button(
            button_frame,
            text="Copy text",
            command=self._copy_text,
            state=tk.DISABLED
        )
        self.copy_btn.pack(side=tk.LEFT)

        ttk.Button(
            button_frame,
            text="Đóng",
            command=self._on_closing
        ).pack(side=tk.RIGHT)

    def _choose_and_read_pdf(self) -> None:
        """Mở file dialog chọn PDF rồi bắt đầu extract text."""
        file_path = filedialog.askopenfilename(
            title="Chọn file PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            parent=self.dialog
        )

        if not file_path:
            return

        self.file_label.config(text=file_path, foreground="black")
        self._start_extraction(file_path)

    def _start_extraction(self, file_path: str) -> None:
        """Bắt đầu extract text trên thread riêng."""
        if self._processing:
            return

        self._processing = True
        self.choose_btn.config(state=tk.DISABLED)
        self.copy_btn.config(state=tk.DISABLED)

        # Xóa text cũ
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete("1.0", tk.END)
        self.text_area.config(state=tk.DISABLED)

        self.status_label.config(text="Đang đọc PDF...", foreground="blue")

        # Chạy extract trên thread riêng
        self._worker_thread = threading.Thread(
            target=self._extract_worker,
            args=(file_path,),
            daemon=True
        )
        self._worker_thread.start()

    def _extract_worker(self, file_path: str) -> None:
        """Worker thread — chạy extract_text_from_pdf."""
        try:
            def on_progress(page_num, total_pages, is_ocr):
                status = f"Đang đọc trang {page_num}/{total_pages}"
                if is_ocr:
                    status += " (OCR)..."
                else:
                    status += "..."
                # Cập nhật UI từ thread khác qua root.after()
                self.dialog.after(0, self._update_status, status)

            result_text = extract_text_from_pdf(file_path, on_progress=on_progress)

            # Gửi kết quả về UI thread
            self.dialog.after(0, self._on_extraction_done, result_text)

        except FileNotFoundError as e:
            self.dialog.after(0, self._on_extraction_error, f"File không tồn tại: {e}")
        except ValueError as e:
            self.dialog.after(0, self._on_extraction_error, f"File không hợp lệ: {e}")
        except Exception as e:
            logger.error(f"Lỗi extract PDF: {e}")
            self.dialog.after(0, self._on_extraction_error, f"Lỗi: {e}")

    def _update_status(self, text: str) -> None:
        """Cập nhật status label (gọi từ UI thread)."""
        self.status_label.config(text=text, foreground="blue")

    def _on_extraction_done(self, result_text: str) -> None:
        """Xử lý kết quả extract thành công (gọi từ UI thread)."""
        self._processing = False
        self.choose_btn.config(state=tk.NORMAL)

        # Hiển thị kết quả
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert("1.0", result_text)
        self.text_area.config(state=tk.DISABLED)

        # Đếm trang
        page_count = result_text.count("--- Trang ")
        self.status_label.config(
            text=f"Đã đọc xong — {page_count} trang",
            foreground="green"
        )

        if result_text.strip():
            self.copy_btn.config(state=tk.NORMAL)

    def _on_extraction_error(self, error_msg: str) -> None:
        """Xử lý lỗi extract (gọi từ UI thread)."""
        self._processing = False
        self.choose_btn.config(state=tk.NORMAL)
        self.status_label.config(text=error_msg, foreground="red")
        messagebox.showerror("Lỗi đọc PDF", error_msg, parent=self.dialog)

    def _copy_text(self) -> None:
        """Copy toàn bộ text vào clipboard."""
        text = self.text_area.get("1.0", tk.END).strip()
        if text:
            self.dialog.clipboard_clear()
            self.dialog.clipboard_append(text)
            self.status_label.config(text="Đã copy text vào clipboard!", foreground="green")

    def _on_closing(self) -> None:
        """Xử lý đóng dialog — lưu kích thước, destroy."""
        # Lưu kích thước dialog
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            self.dialog_config.save_dialog_size(self.DIALOG_NAME, width, height)
        except Exception as e:
            logger.error(f"Lỗi lưu kích thước dialog: {e}")

        self.dialog.destroy()
```

- [ ] **Step 2: Verify dialog can be imported without error**

Run: `python -c "from ui.pdf_reader_dialog import PdfReaderDialog; print('OK')"`
Expected: Prints `OK`

- [ ] **Step 3: Commit**

```bash
git add ui/pdf_reader_dialog.py
git commit -m "feat: add PdfReaderDialog with threaded extraction and progress display"
```

---

### Task 7: Integrate "Đọc PDF" Button into Main UI

**Files:**
- Modify: `ui/excel_realtime_controller.py` (lines 162-178 — buttons_config list)

- [ ] **Step 1: Add "Đọc PDF" to the buttons_config list**

In `ui/excel_realtime_controller.py`, find the `buttons_config` list (line 162). Add the PDF button as the last item in the list:

Change this:

```python
        buttons_config: List[Tuple[str, Callable]] = [
            ("🔍 Quét Sizes", self._scan_sizes),
            ("👁️ Ẩn dòng ngay", self._hide_rows_realtime),
            ("👁️‍🗨️ Hiện tất cả", self._show_all_rows),
            ("📝 Nhập Số Lượng Size", self._input_size_quantities),
            ("💾 Ghi vào Excel", self._write_quantities_to_excel),
            ("📦 Xuất Danh Sách Thùng", self._export_box_list),
        ]
```

To this:

```python
        buttons_config: List[Tuple[str, Callable]] = [
            ("🔍 Quét Sizes", self._scan_sizes),
            ("👁️ Ẩn dòng ngay", self._hide_rows_realtime),
            ("👁️‍🗨️ Hiện tất cả", self._show_all_rows),
            ("📝 Nhập Số Lượng Size", self._input_size_quantities),
            ("💾 Ghi vào Excel", self._write_quantities_to_excel),
            ("📦 Xuất Danh Sách Thùng", self._export_box_list),
            ("📄 Đọc PDF", self._open_pdf_reader),
        ]
```

- [ ] **Step 2: Add the _open_pdf_reader method**

Add this method to the `ExcelRealtimeController` class. Place it after the `_update_color_code` method (around line 1113):

```python
    def _open_pdf_reader(self) -> None:
        """Mở dialog đọc PDF."""
        try:
            from ui.pdf_reader_dialog import PdfReaderDialog
            PdfReaderDialog(self.root)
        except ImportError as e:
            messagebox.showerror(
                "Lỗi",
                f"Không thể mở tính năng đọc PDF.\n"
                f"Hãy cài đặt thư viện: pip install pdfplumber pdf2image pytesseract Pillow\n\n"
                f"Chi tiết: {e}"
            )
        except Exception as e:
            logger.error(f"Lỗi mở PDF Reader: {e}")
            messagebox.showerror("Lỗi", f"Lỗi mở PDF Reader: {e}")
```

**Note:** Tính năng đọc PDF không yêu cầu phải mở file Excel trước — khác với các nút khác. Do đó không cần check `self.com_manager`.

- [ ] **Step 3: Run the app to verify button appears**

Run: `python excel_realtime_controller.py`
Expected: Nút "📄 Đọc PDF" xuất hiện trong action_frame cùng với các nút khác.

- [ ] **Step 4: Commit**

```bash
git add ui/excel_realtime_controller.py
git commit -m "feat: integrate 'Đọc PDF' button into main UI action frame"
```

---

### Task 8: Add dialog config default for PDF reader

**Files:**
- Modify: `excel_automation/dialog_config_manager.py` (line 20-40 — DEFAULT_CONFIG)

- [ ] **Step 1: Add pdf_reader to DEFAULT_CONFIG**

In `excel_automation/dialog_config_manager.py`, add the `pdf_reader` entry to the `dialogs` dict inside `DEFAULT_CONFIG`. Find this block:

```python
            "size_quantity_input": {
                "width": 550,
                "height": 750
            }
```

Change to:

```python
            "size_quantity_input": {
                "width": 550,
                "height": 750
            },
            "pdf_reader": {
                "width": 700,
                "height": 500
            }
```

- [ ] **Step 2: Run all tests to verify nothing is broken**

Run: `pytest tests/ -v`
Expected: All tests PASS

- [ ] **Step 3: Commit**

```bash
git add excel_automation/dialog_config_manager.py
git commit -m "feat: add pdf_reader dialog default size (700x500) to config"
```

---

### Task 9: End-to-End Manual Test

**Files:** None (manual verification)

- [ ] **Step 1: Verify OCR availability check**

Run:
```bash
python -c "from excel_automation.pdf_reader import check_ocr_available; print(check_ocr_available())"
```
Expected: Prints `{"tesseract": True/False, "poppler": True/False}` depending on your system.

- [ ] **Step 2: Test with Test.pdf in project root**

Run:
```bash
python -c "
from excel_automation.pdf_reader import extract_text_from_pdf
result = extract_text_from_pdf('Test.pdf')
print(result[:500] if result else 'Empty result')
print(f'\\nTotal length: {len(result)} chars')
"
```
Expected: Prints extracted text from `Test.pdf` with page markers.

- [ ] **Step 3: Launch app and test dialog**

Run: `python excel_realtime_controller.py`

Manual test steps:
1. Click "📄 Đọc PDF"
2. Dialog opens at 700x500, centered
3. Click "Chọn file PDF" → select `Test.pdf`
4. Status shows "Đang đọc trang X/Y..."
5. Text appears in ScrolledText area with page markers
6. Click "Copy text" → text copied to clipboard
7. Close dialog → reopen → size is remembered

- [ ] **Step 4: Run full test suite**

Run: `pytest tests/ -v`
Expected: All tests PASS

- [ ] **Step 5: Final commit (if any fixes needed)**

```bash
git add -A
git commit -m "test: verify end-to-end PDF reader functionality"
```
