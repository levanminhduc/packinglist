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
