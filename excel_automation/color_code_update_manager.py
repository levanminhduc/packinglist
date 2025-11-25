from typing import Tuple
import logging

logger = logging.getLogger(__name__)


class ColorCodeUpdateManager:
    
    def __init__(self, config):
        self.config = config
    
    def get_current_color_code(self, worksheet, column: str = 'E') -> str:
        start_row = self.config.get_start_row()
        col_num = self._column_to_number(column)
        
        try:
            cell_value = worksheet.Cells(start_row, col_num).Value
            if cell_value is not None:
                value_str = str(cell_value)
                if value_str.startswith("'"):
                    return value_str[1:]
                return value_str
            return ""
        except Exception as e:
            logger.error(f"Lỗi khi đọc mã màu: {e}")
            return ""
    
    def update_color_code_bulk(self, worksheet, new_color_code: str, column: str = 'E') -> int:
        start_row = self.config.get_start_row()
        end_row = self.config.get_end_row()
        col_num = self._column_to_number(column)
        
        updated_count = 0
        
        try:
            prefixed_value = f"'{new_color_code}"
            
            for row in range(start_row, end_row + 1):
                worksheet.Cells(row, col_num).Value = prefixed_value
                updated_count += 1
            
            logger.info(f"Đã cập nhật {updated_count} dòng mã màu thành '{prefixed_value}'")
            return updated_count
            
        except Exception as e:
            logger.error(f"Lỗi khi cập nhật mã màu: {e}")
            raise RuntimeError(f"Không thể cập nhật mã màu: {str(e)}")
    
    def validate_color_code(self, color_code: str) -> Tuple[bool, str]:
        if not color_code or not color_code.strip():
            return False, "Mã màu không được để trống"
        
        return True, ""
    
    def _column_to_number(self, column: str) -> int:
        column = column.upper()
        result = 0
        for char in column:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

