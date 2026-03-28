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

    @staticmethod
    def _normalize_size(size_str: str) -> str:
        size_str = size_str.strip()
        try:
            size_num = int(size_str)
            return str(size_num).zfill(3)
        except ValueError:
            return size_str

    @staticmethod
    def _extract_size_quantities(full_text: str) -> Dict[str, int]:
        lines = full_text.split('\n')
        size_quantities: Dict[str, int] = {}

        current_qty = None
        for line in lines:
            line = line.strip()

            qty_match = re.match(
                r'^0{2,}\d+\s+\d{11,}\s+.*?\s+(\d+)\s+[\d.]+\s+[\d,.]+\s+\w{3}$',
                line
            )
            if qty_match:
                current_qty = int(qty_match.group(1))
                continue

            size_match = re.match(r'^Size:\s*(.+)$', line)
            if size_match and current_qty is not None:
                raw_size = size_match.group(1).strip()
                normalized = PDFPOParser._normalize_size(raw_size)
                size_quantities[normalized] = current_qty
                current_qty = None

        if not size_quantities:
            raise RuntimeError("Không tìm thấy dữ liệu Size/Quantity trong file PDF")

        return size_quantities
