import re
from typing import Tuple, Optional
import logging

logger = logging.getLogger(__name__)


class POUpdateManager:
    
    def __init__(self, config):
        self.config = config
    
    def get_current_po(self, worksheet, column: str = 'A') -> str:
        start_row = self.config.get_start_row()
        col_num = self._column_to_number(column)

        try:
            cell_value = worksheet.Cells(start_row, col_num).Value
            if cell_value is not None:
                value_str = str(cell_value)
                if value_str.endswith('.0'):
                    return value_str[:-2]
                return value_str
            return ""
        except Exception as e:
            logger.error(f"Lỗi khi đọc PO: {e}")
            return ""
    
    def update_po_bulk(self, worksheet, new_po: str, column: str = 'A') -> int:
        start_row = self.config.get_start_row()
        end_row = self.config.get_end_row()
        col_num = self._column_to_number(column)
        
        updated_count = 0
        
        try:
            for row in range(start_row, end_row + 1):
                worksheet.Cells(row, col_num).Value = new_po
                updated_count += 1
            
            logger.info(f"Đã cập nhật {updated_count} dòng PO thành '{new_po}'")
            return updated_count
            
        except Exception as e:
            logger.error(f"Lỗi khi cập nhật PO: {e}")
            raise RuntimeError(f"Không thể cập nhật PO: {str(e)}")
    
    def validate_po(self, po_value: str) -> Tuple[bool, str]:
        if not po_value or not po_value.strip():
            return False, "Giá trị không được để trống"

        return True, ""
    
    def _column_to_number(self, column: str) -> int:
        column = column.upper()
        result = 0
        for char in column:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

