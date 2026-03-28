from dataclasses import dataclass, field
from typing import Dict
import logging
import re

logger = logging.getLogger(__name__)


@dataclass
class PDFPOData:
    raw_po: str
    po_number: str
    color_code: str
    size_quantities: Dict[str, int] = field(default_factory=dict)
    total_quantity: int = 0
    source_file: str = ""


class PDFPOParser:

    @staticmethod
    def _extract_po_number(full_text: str) -> tuple:
        pattern = r'P\.?\s*O\.?\s*No\.?\s*[^\n]*\n\s*(\S+)'
        match = re.search(pattern, full_text)
        if not match:
            raise RuntimeError("Không tìm thấy PO Number trong file PDF")

        raw_po = match.group(1).strip()
        po_part = raw_po.split('-')[0]
        cleaned = po_part.lstrip('0') or '0'
        return (raw_po, cleaned)

    @staticmethod
    def _extract_color_code(full_text: str) -> str:
        pattern = r'(?:^|\s)(\d{11,})\s'
        match = re.search(pattern, full_text)
        if not match:
            raise RuntimeError("Không tìm thấy Article Number trong file PDF")

        article_no = match.group(1)
        color_code = article_no[4:8]
        return color_code
