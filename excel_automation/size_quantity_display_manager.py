from typing import Dict, List, Optional, Tuple
from win32com.client import CDispatch
import logging

from excel_automation.size_filter_config import SizeFilterConfig
from excel_automation.carton_allocation_calculator import AllocationResult
from excel_automation.utils import get_size_sort_key

logger = logging.getLogger(__name__)


class SizeQuantityDisplayManager:
    
    def __init__(self, config: SizeFilterConfig):
        self.config = config
    
    def _get_size_row_mapping(
        self,
        worksheet: CDispatch,
        column: str,
        start_row: int,
        end_row: int
    ) -> Dict[str, List[int]]:
        col_num = self._column_letter_to_number(column)
        size_rows: Dict[str, List[int]] = {}
        
        for row in range(start_row, end_row + 1):
            cell_value = worksheet.Cells(row, col_num).Value
            
            if cell_value is not None:
                size_str = str(cell_value).strip()
                if size_str.isdigit():
                    size_str = size_str.zfill(3)
                
                if size_str:
                    if size_str not in size_rows:
                        size_rows[size_str] = []
                    size_rows[size_str].append(row)
        
        logger.info(f"Đã map {len(size_rows)} sizes với các dòng tương ứng")
        return size_rows
    
    def write_quantities_to_excel(
        self,
        excel_app: CDispatch,
        worksheet: CDispatch,
        selected_sizes: List[str],
        size_quantities: Dict[str, Optional[int]],
        current_quantities: Dict[str, Optional[int]],
        size_column: str,
        start_row: Optional[int] = None,
        end_row: Optional[int] = None
    ) -> int:
        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.config.get_end_row()

        size_row_mapping = self._get_size_row_mapping(
            worksheet,
            size_column,
            start_row,
            end_row
        )

        written_count = 0

        try:
            excel_app.ScreenUpdating = False

            for position, size in enumerate(selected_sizes, start=1):
                column_number = 6 + position

                if size not in size_row_mapping:
                    logger.warning(f"Size {size} không tìm thấy trong mapping")
                    continue

                row_number = size_row_mapping[size][0]

                if size in size_quantities:
                    quantity = size_quantities[size]

                    if quantity is not None:
                        worksheet.Cells(row_number, column_number).Value = quantity
                        logger.info(
                            f"Đã ghi size {size}: {quantity} thùng vào cell "
                            f"({row_number}, {column_number})"
                        )
                        written_count += 1
                    elif size in current_quantities and current_quantities[size] is not None:
                        worksheet.Cells(row_number, column_number).Value = None
                        logger.info(
                            f"Đã xóa size {size} tại cell ({row_number}, {column_number})"
                        )

            logger.info(f"Đã ghi {written_count} cells thành công")
            return written_count

        except Exception as e:
            logger.error(f"Lỗi khi ghi số lượng vào Excel: {e}", exc_info=True)
            raise
        finally:
            excel_app.ScreenUpdating = True
    
    def get_current_quantities(
        self,
        worksheet: CDispatch,
        selected_sizes: List[str],
        size_column: str,
        start_row: Optional[int] = None,
        end_row: Optional[int] = None
    ) -> Dict[str, Optional[int]]:
        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.config.get_end_row()

        size_row_mapping = self._get_size_row_mapping(
            worksheet,
            size_column,
            start_row,
            end_row
        )

        current_quantities: Dict[str, Optional[int]] = {}

        for position, size in enumerate(selected_sizes, start=1):
            column_number = 6 + position

            if size in size_row_mapping:
                row_number = size_row_mapping[size][0]

                try:
                    cell_value = worksheet.Cells(row_number, column_number).Value

                    if cell_value is not None:
                        try:
                            quantity = int(cell_value)
                            current_quantities[size] = quantity
                            logger.info(
                                f"Đọc size {size}: {quantity} thùng từ cell "
                                f"({row_number}, {column_number})"
                            )
                        except (ValueError, TypeError):
                            current_quantities[size] = None
                            logger.warning(
                                f"Size {size} tại cell ({row_number}, {column_number}) "
                                f"có giá trị không hợp lệ: {cell_value}"
                            )
                    else:
                        current_quantities[size] = None
                except Exception as e:
                    logger.warning(
                        f"Lỗi khi đọc cell ({row_number}, {column_number}) "
                        f"cho size {size}: {e}"
                    )
                    current_quantities[size] = None
            else:
                logger.warning(f"Size {size} không tìm thấy trong mapping")
                current_quantities[size] = None

        logger.info(f"Đã đọc số lượng hiện tại cho {len(current_quantities)} sizes")
        return current_quantities

    def _column_letter_to_number(self, column: str) -> int:
        column = column.upper()
        result = 0
        for char in column:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    def _has_formula(self, worksheet: CDispatch, row: int, column: int) -> bool:
        try:
            cell = worksheet.Cells(row, column)
            if cell.HasFormula:
                return True
            return False
        except Exception:
            return False

    def write_allocated_quantities_to_excel(
        self,
        excel_app: CDispatch,
        worksheet: CDispatch,
        allocation_result: AllocationResult,
        selected_sizes: List[str],
        size_column: str,
        start_row: Optional[int] = None,
        end_row: Optional[int] = None,
        box_start_row: int = 15,
        box_end_row: int = 16
    ) -> Tuple[int, int]:
        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.config.get_end_row()

        size_row_mapping = self._get_size_row_mapping(
            worksheet,
            size_column,
            start_row,
            end_row
        )

        sorted_sizes = sorted(selected_sizes, key=get_size_sort_key)

        column_assignments: List[Dict] = []
        current_box = 1
        current_column = 7

        for size in sorted_sizes:
            if size not in allocation_result.allocations:
                continue

            alloc = allocation_result.allocations[size]
            if alloc.full_boxes > 0:
                box_start = current_box
                box_end = current_box + alloc.full_boxes - 1

                column_assignments.append({
                    'type': 'full',
                    'size': size,
                    'column': current_column,
                    'quantity': alloc.full_qty,
                    'box_start': box_start,
                    'box_end': box_end
                })

                current_box = box_end + 1
                current_column += 1

        for i, carton in enumerate(allocation_result.combined_cartons):
            box_start = current_box
            box_end = current_box

            for size in carton.sizes:
                if size in size_row_mapping:
                    column_assignments.append({
                        'type': 'combined',
                        'size': size,
                        'column': current_column,
                        'quantity': carton.quantities[size],
                        'box_start': box_start,
                        'box_end': box_end,
                        'carton_index': i
                    })

            current_box += 1
            current_column += 1

        written_count = 0
        columns_used = 0

        try:
            excel_app.ScreenUpdating = False

            processed_columns = set()

            for assignment in column_assignments:
                size = assignment['size']
                column_number = assignment['column']
                quantity = assignment['quantity']
                box_start = assignment['box_start']
                box_end = assignment['box_end']

                if size not in size_row_mapping:
                    logger.warning(f"Size {size} khong tim thay trong mapping")
                    continue

                row_number = size_row_mapping[size][0]

                worksheet.Cells(row_number, column_number).Value = quantity
                logger.info(
                    f"Da ghi size {size}: {quantity} pcs vao cell "
                    f"({row_number}, {column_number})"
                )
                written_count += 1

                if column_number not in processed_columns:
                    start_has_formula = self._has_formula(
                        worksheet, box_start_row, column_number
                    )
                    end_has_formula = self._has_formula(
                        worksheet, box_end_row, column_number
                    )

                    if not start_has_formula:
                        worksheet.Cells(box_start_row, column_number).Value = box_start
                    else:
                        logger.info(
                            f"Bo qua ghi dong {box_start_row} cot {column_number} "
                            f"vi co cong thuc"
                        )

                    if not end_has_formula:
                        worksheet.Cells(box_end_row, column_number).Value = box_end
                    else:
                        logger.info(
                            f"Bo qua ghi dong {box_end_row} cot {column_number} "
                            f"vi co cong thuc"
                        )

                    if not start_has_formula or not end_has_formula:
                        logger.info(
                            f"Da ghi From/To Ctn: {box_start}-{box_end} vao cot {column_number}"
                        )

                    processed_columns.add(column_number)
                    columns_used += 1

            logger.info(
                f"Da ghi {written_count} cells, {columns_used} cot, "
                f"tong {allocation_result.total_boxes} thung"
            )
            return written_count, columns_used

        except Exception as e:
            logger.error(f"Loi khi ghi so luong vao Excel: {e}", exc_info=True)
            raise
        finally:
            excel_app.ScreenUpdating = True
