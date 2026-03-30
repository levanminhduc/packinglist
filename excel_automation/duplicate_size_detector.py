from typing import Dict, List, Optional
import logging

from excel_automation.utils import normalize_size_value

logger = logging.getLogger(__name__)


class DuplicateSizeDetector:

    def __init__(self, com_manager):
        self.com_manager = com_manager

    def detect_duplicates(
        self,
        column: Optional[str] = None,
        start_row: Optional[int] = None,
        end_row: Optional[int] = None
    ) -> Dict[str, List[int]]:
        worksheet = self.com_manager.worksheet
        if worksheet is None:
            return {}

        column = column or self.com_manager.config.get_column()
        start_row = start_row or self.com_manager.config.get_start_row()
        end_row = end_row or self.com_manager.detect_end_row()

        try:
            range_str = f"{column}{start_row}:{column}{end_row}"
            raw_values = worksheet.Range(range_str).Value

            if raw_values is None:
                return {}

            if not isinstance(raw_values, tuple):
                raw_values = ((raw_values,),)

            size_rows: Dict[str, List[int]] = {}

            for row_offset, row_tuple in enumerate(raw_values):
                cell_value = row_tuple[0] if isinstance(row_tuple, tuple) else row_tuple

                if cell_value is not None:
                    size_str = normalize_size_value(cell_value)
                    if size_str:
                        actual_row = start_row + row_offset
                        if size_str not in size_rows:
                            size_rows[size_str] = []
                        size_rows[size_str].append(actual_row)

            duplicates = {
                size: rows
                for size, rows in size_rows.items()
                if len(rows) >= 2
            }

            if duplicates:
                logger.info(
                    f"Phát hiện {len(duplicates)} size trùng: "
                    f"{', '.join(f'{s}({len(r)} dòng)' for s, r in duplicates.items())}"
                )
            else:
                logger.info("Không phát hiện size trùng")

            return duplicates

        except Exception as e:
            logger.error(f"Lỗi khi detect size trùng: {e}")
            raise RuntimeError(f"Không thể quét size trùng: {str(e)}")
