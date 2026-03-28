"""
PDF Reader Module — Extract text từ PDF files.

Hỗ trợ:
- Digital PDF: extract text trực tiếp bằng pdfplumber
- Scanned PDF: convert trang → ảnh rồi OCR bằng pytesseract

Module này không phụ thuộc UI — chỉ chứa logic thuần.
"""

import logging
import os
import shutil
from typing import Callable, Optional

logger = logging.getLogger(__name__)

MIN_TEXT_LENGTH = 10  # Ngưỡng ký tự tối thiểu để coi là trang digital


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
        pass

    # Check Poppler (pdf2image cần pdftoppm binary)
    if shutil.which("pdftoppm") is not None:
        result["poppler"] = True

    return result
