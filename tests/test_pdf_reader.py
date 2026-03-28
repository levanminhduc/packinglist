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
        assert "Trang 3" in result or "trang 3" in result


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


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
