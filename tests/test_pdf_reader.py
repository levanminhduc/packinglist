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


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
