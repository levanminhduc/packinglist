from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Optional
import logging
import re

import pdfplumber

logger = logging.getLogger(__name__)


@dataclass
class PDFPOData:
    raw_po: str
    po_number: str
    color_code: str
    size_quantities: Dict[str, int] = field(default_factory=dict)
    total_quantity: int = 0
    source_file: str = ""
    ordertotal_from_pdf: Optional[int] = None
    quantity_mismatch: bool = False


class PDFPOParser:

    @staticmethod
    def _extract_po_number(full_text: str) -> tuple:
        pattern = r'(\d{7,}-\d+)\s'
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

    @staticmethod
    def _extract_ordertotal(full_text: str) -> Optional[int]:
        pattern = r'Ordertotal\s+(\d+)\s+'
        match = re.search(pattern, full_text)
        if not match:
            logger.warning("Không tìm thấy dòng Ordertotal trong PDF")
            return None
        return int(match.group(1))

    @staticmethod
    def parse(file_path: str) -> "PDFPOData":
        path = Path(file_path)
        if not path.exists():
            raise RuntimeError(f"Không thể đọc file PDF: File không tồn tại: {file_path}")

        try:
            full_text = ""
            with pdfplumber.open(str(path)) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        full_text += page_text + "\n"
        except Exception as e:
            logger.error(f"Lỗi khi đọc file PDF: {e}")
            raise RuntimeError(f"Không thể đọc file PDF: {str(e)}")

        if not full_text.strip():
            raise RuntimeError("File PDF không chứa nội dung text")

        raw_po, po_number = PDFPOParser._extract_po_number(full_text)
        color_code = PDFPOParser._extract_color_code(full_text)
        size_quantities = PDFPOParser._extract_size_quantities(full_text)
        total_quantity = sum(size_quantities.values())

        ordertotal = PDFPOParser._extract_ordertotal(full_text)
        mismatch = ordertotal is not None and total_quantity != ordertotal
        if mismatch:
            logger.warning(
                f"Chênh lệch qty! Parse={total_quantity}, Ordertotal PDF={ordertotal}, "
                f"thiếu={ordertotal - total_quantity}"
            )

        logger.info(
            f"Parse PDF thành công: PO={po_number}, Color={color_code}, "
            f"{len(size_quantities)} sizes, total={total_quantity}"
        )

        return PDFPOData(
            raw_po=raw_po,
            po_number=po_number,
            color_code=color_code,
            size_quantities=size_quantities,
            total_quantity=total_quantity,
            source_file=str(path),
            ordertotal_from_pdf=ordertotal,
            quantity_mismatch=mismatch
        )
