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


def is_scanned_page(page) -> bool:
    """
    Kiểm tra trang PDF có phải là trang scan (ảnh) hay không.

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
        page_number: Số thứ tự trang (1-based)

    Returns:
        Text extract được từ trang
    """
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

        images = convert_from_path(
            page.pdf.stream.name,
            first_page=page_number,
            last_page=page_number,
            dpi=300
        )

        if not images:
            return f"[Trang {page_number}: Không thể convert trang sang ảnh]"

        text = pytesseract.image_to_string(images[0], lang="vie")
        return text.strip() if text else f"[Trang {page_number}: OCR không detect được text]"

    except Exception as e:
        logger.error(f"Lỗi OCR trang {page_number}: {e}")
        return f"[Trang {page_number}: Lỗi khi OCR — {e}]"


def extract_text_from_pdf(
    file_path: str,
    on_progress: Optional[Callable] = None
) -> str:
    """
    Extract toàn bộ text từ file PDF.

    Args:
        file_path: Đường dẫn file PDF
        on_progress: Callback(page_num, total_pages, is_ocr)

    Returns:
        Text gộp từ tất cả trang PDF

    Raises:
        FileNotFoundError: File không tồn tại
        ValueError: File không phải PDF
        RuntimeError: PDF bị password hoặc lỗi khác
    """
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

                if on_progress:
                    on_progress(page_number, total_pages, is_ocr)

                page_text = extract_page_text(page, page_number)

                all_text_parts.append(f"--- Trang {page_number} ---")
                all_text_parts.append(page_text if page_text else "")

        return "\n".join(all_text_parts)

    except Exception as e:
        if isinstance(e, (FileNotFoundError, ValueError)):
            raise
        # Check for password/encryption errors — walk the full exception chain.
        # pdfplumber wraps PDFPasswordIncorrect in PdfminerException via __context__
        # (not __cause__), and both have empty str() representations.
        def _is_password_exc(exc) -> bool:
            seen = set()
            current = exc
            while current is not None and id(current) not in seen:
                seen.add(id(current))
                name = type(current).__name__.lower()
                msg = str(current).lower()
                if (
                    "password" in name or "incorrect" in name
                    or "password" in msg or "encrypted" in msg
                ):
                    return True
                current = current.__cause__ or current.__context__
            return False

        if _is_password_exc(e):
            raise RuntimeError(
                f"File PDF được bảo vệ bằng mật khẩu (password). "
                f"Vui lòng mở khóa file trước khi đọc."
            ) from e
        logger.error(f"Lỗi khi đọc PDF: {e}")
        raise RuntimeError(f"Lỗi khi đọc PDF: {e}") from e


