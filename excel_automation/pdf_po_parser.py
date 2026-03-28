from dataclasses import dataclass, field
from typing import Dict
import logging

logger = logging.getLogger(__name__)


@dataclass
class PDFPOData:
    raw_po: str
    po_number: str
    color_code: str
    size_quantities: Dict[str, int] = field(default_factory=dict)
    total_quantity: int = 0
    source_file: str = ""
