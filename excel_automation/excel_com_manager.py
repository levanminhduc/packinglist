from typing import List, Optional, Set
from pathlib import Path
import logging
import win32com.client
from win32com.client import CDispatch

from excel_automation.size_filter_config import SizeFilterConfig
from excel_automation.utils import get_size_sort_key, normalize_size_value, find_last_data_row

logger = logging.getLogger(__name__)


class ExcelCOMManager:
    
    def __init__(self, config: Optional[SizeFilterConfig] = None):
        self.config = config or SizeFilterConfig()
        self.excel_app: Optional[CDispatch] = None
        self.workbook: Optional[CDispatch] = None
        self.worksheet: Optional[CDispatch] = None
        self.current_file: Optional[str] = None
        self.current_sheet: Optional[str] = None
        
        logger.info("Khởi tạo ExcelCOMManager")
    
    def _is_excel_alive(self) -> bool:
        if self.excel_app is None:
            return False
        try:
            _ = self.excel_app.Version
            return True
        except Exception:
            logger.warning("Excel COM reference bị stale, sẽ tạo lại")
            self.excel_app = None
            self.workbook = None
            self.worksheet = None
            return False

    def _init_excel_app(self) -> None:
        try:
            self.excel_app = win32com.client.Dispatch("Excel.Application")
            self.excel_app.Visible = True
        except Exception as e:
            logger.warning(f"Lỗi khi khởi tạo Excel.Application: {e}, thử lại...")
            try:
                import pythoncom
                pythoncom.CoInitialize()
                self.excel_app = win32com.client.DispatchEx("Excel.Application")
                self.excel_app.Visible = True
            except Exception as e2:
                logger.error(f"Không thể khởi tạo Excel: {e2}")
                raise RuntimeError("Không thể mở Excel. Vui lòng kiểm tra:\n- File có tồn tại không\n- Excel có đang mở file này không\n- Bạn có quyền truy cập file không")

        self.excel_app.DisplayAlerts = False
        logger.info("Đã khởi tạo Excel Application")

    def open_excel_file(self, file_path: str) -> None:
        file_path_obj = Path(file_path)
        if not file_path_obj.exists():
            raise FileNotFoundError(f"File không tồn tại: {file_path}")

        try:
            if not self._is_excel_alive():
                self._init_excel_app()

            if self.workbook is not None:
                try:
                    self.workbook.Close(SaveChanges=False)
                    logger.info("Đã đóng workbook cũ trước khi mở file mới")
                except Exception:
                    pass
                self.workbook = None
                self.worksheet = None

            abs_path = str(file_path_obj.absolute())
            logger.info(f"Đang mở workbook: {abs_path}")
            self.workbook = self.excel_app.Workbooks.Open(abs_path)
            self.current_file = file_path

            wb_count = self.excel_app.Workbooks.Count
            active_wb = self.excel_app.ActiveWorkbook
            active_name = active_wb.Name if active_wb else "None"
            logger.info(f"Workbooks đang mở: {wb_count}, ActiveWorkbook: {active_name}, Opened: {self.workbook.Name}")

            sheet_name = self.config.get_sheet_name()
            try:
                self.worksheet = self.workbook.Sheets(sheet_name)
                self.current_sheet = sheet_name
            except Exception:
                self.worksheet = self.workbook.Sheets(1)
                self.current_sheet = self.worksheet.Name
                logger.warning(f"Sheet '{sheet_name}' không tồn tại, sử dụng sheet đầu tiên: {self.current_sheet}")

            self.workbook.Activate()
            self.worksheet.Activate()
            self.excel_app.Visible = True
            logger.info(f"Đã mở file: {file_path}, sheet: {self.current_sheet}")

        except Exception as e:
            logger.error(f"Lỗi khi mở file Excel qua COM: {e}")
            self._cleanup_on_error()
            raise RuntimeError(f"Không thể mở file Excel: {str(e)}")
    
    def get_sheet_names(self) -> List[str]:
        if self.workbook is None:
            raise RuntimeError("Chưa mở workbook nào")
        
        try:
            sheet_names = []
            for sheet in self.workbook.Sheets:
                sheet_names.append(sheet.Name)
            
            logger.info(f"Lấy được {len(sheet_names)} sheets")
            return sheet_names
            
        except Exception as e:
            logger.error(f"Lỗi khi lấy danh sách sheets: {e}")
            raise RuntimeError(f"Không thể lấy danh sách sheets: {str(e)}")
    
    def switch_sheet(self, sheet_name: str) -> None:
        if self.workbook is None:
            raise RuntimeError("Chưa mở workbook nào")
        
        try:
            self.worksheet = self.workbook.Sheets(sheet_name)
            self.worksheet.Activate()
            self.current_sheet = sheet_name
            logger.info(f"Đã chuyển sang sheet: {sheet_name}")
            
        except Exception as e:
            logger.error(f"Lỗi khi chuyển sheet: {e}")
            raise RuntimeError(f"Không thể chuyển sang sheet '{sheet_name}': {str(e)}")
    
    def detect_end_row(self, reference_column: str = 'A') -> int:
        """Tự nhận diện dòng cuối cùng có dữ liệu trong worksheet."""
        if self.worksheet is None:
            return self.config.get_end_row()

        col_num = self._column_letter_to_number(reference_column)
        start_row = self.config.get_start_row()
        detected = find_last_data_row(self.worksheet, col_num, start_row)
        return detected

    def scan_sizes(self, column: Optional[str] = None, start_row: Optional[int] = None,
                   end_row: Optional[int] = None) -> List[str]:
        if self.worksheet is None:
            raise RuntimeError("Chưa chọn worksheet nào")

        column = column or self.config.get_column()
        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.detect_end_row()

        try:
            sizes: Set[str] = set()
            col_num = self._column_letter_to_number(column)

            for row in range(start_row, end_row + 1):
                cell_value = self.worksheet.Cells(row, col_num).Value

                if cell_value is not None:
                    size_str = normalize_size_value(cell_value)

                    if size_str:
                        # Nếu giá trị gốc là số lẻ → ghi giá trị đã làm tròn lại vào Excel
                        self._fix_decimal_cell(row, col_num, cell_value, size_str)
                        sizes.add(size_str)

            sorted_sizes = sorted(sizes, key=get_size_sort_key)
            logger.info(f"Quét được {len(sorted_sizes)} size khác nhau trong {column}[{start_row}:{end_row}]")
            return sorted_sizes

        except Exception as e:
            logger.error(f"Lỗi khi quét sizes: {e}")
            raise RuntimeError(f"Không thể quét sizes: {str(e)}")

    def _fix_decimal_cell(self, row: int, col_num: int, original_value, normalized: str) -> None:
        """Ghi lại giá trị đã làm tròn lên vào Excel nếu cell gốc là số lẻ."""
        try:
            needs_fix = False

            if isinstance(original_value, float) and original_value != int(original_value):
                needs_fix = True
            elif isinstance(original_value, str):
                raw = original_value.strip()
                check = raw.replace(',', '.') if ',' in raw and '.' not in raw else raw
                try:
                    num = float(check)
                    if num != int(num):
                        needs_fix = True
                except (ValueError, TypeError):
                    pass

            if needs_fix:
                rounded = int(normalized)  # "008" → 8
                self.worksheet.Cells(row, col_num).Value = rounded
                logger.info(
                    f"Đã làm tròn size lẻ: {original_value} → {rounded} "
                    f"tại cell ({row}, {col_num})"
                )
        except Exception as e:
            logger.warning(f"Không thể ghi lại cell ({row}, {col_num}): {e}")
    
    def hide_rows_realtime(self, selected_sizes: List[str], column: Optional[str] = None,
                          start_row: Optional[int] = None, end_row: Optional[int] = None) -> int:
        if self.worksheet is None:
            raise RuntimeError("Chưa chọn worksheet nào")
        
        column = column or self.config.get_column()
        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.detect_end_row()

        try:
            if self.excel_app:
                self.excel_app.ScreenUpdating = False

            selected_set = set(selected_sizes)
            hidden_count = 0
            col_num = self._column_letter_to_number(column)
            
            for row in range(start_row, end_row + 1):
                cell_value = self.worksheet.Cells(row, col_num).Value
                
                if cell_value is not None:
                    size_str = normalize_size_value(cell_value)
                    
                    if size_str not in selected_set:
                        self.worksheet.Rows(row).Hidden = True
                        hidden_count += 1
                    else:
                        self.worksheet.Rows(row).Hidden = False
                else:
                    self.worksheet.Rows(row).Hidden = True
                    hidden_count += 1
            
            if self.excel_app:
                self.excel_app.ScreenUpdating = True
            
            logger.info(f"Đã ẩn {hidden_count} dòng real-time")
            return hidden_count
            
        except Exception as e:
            if self.excel_app:
                self.excel_app.ScreenUpdating = True
            logger.error(f"Lỗi khi ẩn dòng: {e}")
            raise RuntimeError(f"Không thể ẩn dòng: {str(e)}")
    
    def show_all_rows(self, start_row: Optional[int] = None, end_row: Optional[int] = None) -> None:
        if self.worksheet is None:
            raise RuntimeError("Chưa chọn worksheet nào")
        
        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.detect_end_row()

        try:
            if self.excel_app:
                self.excel_app.ScreenUpdating = False

            for row in range(start_row, end_row + 1):
                self.worksheet.Rows(row).Hidden = False

            if self.excel_app:
                self.excel_app.ScreenUpdating = True

            logger.info(f"Đã hiện tất cả dòng từ {start_row} đến {end_row}")
            
        except Exception as e:
            if self.excel_app:
                self.excel_app.ScreenUpdating = True
            logger.error(f"Lỗi khi hiện dòng: {e}")
            raise RuntimeError(f"Không thể hiện dòng: {str(e)}")
    
    def copy_sheet(self, sheet_name: Optional[str] = None) -> str:
        if self.workbook is None:
            raise RuntimeError("Chưa mở workbook nào")

        try:
            source_sheet = self.workbook.Sheets(sheet_name) if sheet_name else self.worksheet
            last_sheet = self.workbook.Sheets(self.workbook.Sheets.Count)
            source_sheet.Copy(None, last_sheet)

            new_sheet = self.workbook.Sheets(self.workbook.Sheets.Count)
            new_name = new_sheet.Name

            logger.info(f"Đã copy sheet '{source_sheet.Name}' → '{new_name}' (cuối workbook)")
            return new_name

        except Exception as e:
            logger.error(f"Lỗi khi copy sheet: {e}")
            raise RuntimeError(f"Không thể copy sheet: {str(e)}")

    def rename_sheet(self, old_name: str, new_name: str) -> None:
        if self.workbook is None:
            raise RuntimeError("Chưa mở workbook nào")

        if not new_name or not new_name.strip():
            raise ValueError("Tên sheet không được rỗng")

        new_name = new_name.strip()

        for sheet in self.workbook.Sheets:
            if sheet.Name.lower() == new_name.lower():
                raise ValueError(f"Sheet '{new_name}' đã tồn tại")

        try:
            self.workbook.Sheets(old_name).Name = new_name
            if self.current_sheet == old_name:
                self.current_sheet = new_name
            logger.info(f"Đã đổi tên sheet '{old_name}' → '{new_name}'")

        except Exception as e:
            logger.error(f"Lỗi khi đổi tên sheet: {e}")
            raise RuntimeError(f"Không thể đổi tên sheet: {str(e)}")

    def clear_quantity_columns(self, start_row: Optional[int] = None,
                               end_row: Optional[int] = None,
                               start_col: int = 7,
                               end_col: int = 39) -> int:
        if self.worksheet is None:
            raise RuntimeError("Chưa chọn worksheet nào")

        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.detect_end_row()

        cleared_count = 0

        try:
            if self.excel_app:
                self.excel_app.ScreenUpdating = False

            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    cell_value = self.worksheet.Cells(row, col).Value
                    if cell_value is not None:
                        self.worksheet.Cells(row, col).Value = None
                        cleared_count += 1

            if self.excel_app:
                self.excel_app.ScreenUpdating = True

            logger.info(f"Đã xóa {cleared_count} ô số lượng (cột {start_col}-{end_col}, dòng {start_row}-{end_row})")
            return cleared_count

        except Exception as e:
            if self.excel_app:
                self.excel_app.ScreenUpdating = True
            logger.error(f"Lỗi khi xóa số lượng: {e}")
            raise RuntimeError(f"Không thể xóa số lượng: {str(e)}")

    def _column_letter_to_number(self, column: str) -> int:
        column = column.upper()
        result = 0
        for char in column:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result
    
    def _number_to_column_letter(self, col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result

    def _cleanup_on_error(self) -> None:
        try:
            if self.workbook:
                self.workbook.Close(SaveChanges=False)
        except Exception:
            pass
        
        try:
            if self.excel_app:
                self.excel_app.Quit()
        except Exception:
            pass
        
        self.workbook = None
        self.worksheet = None
        self.excel_app = None
    
    def detach(self, save_changes: bool = False) -> None:
        try:
            if self.workbook:
                if save_changes:
                    self.workbook.Save()
                    logger.info("Đã lưu workbook")

                self.workbook = None
                self.worksheet = None
                logger.info("Đã detach khỏi workbook (Excel vẫn mở)")

            if self.excel_app:
                self.excel_app = None
                logger.info("Đã detach khỏi Excel Application (Excel vẫn chạy)")

        except Exception as e:
            logger.error(f"Lỗi khi detach COM objects: {e}")
            self.workbook = None
            self.worksheet = None
            self.excel_app = None

    def close(self, save_changes: bool = False) -> None:
        try:
            if self.workbook:
                self.workbook.Close(SaveChanges=save_changes)
                logger.info(f"Đã đóng workbook (save={save_changes})")
                self.workbook = None
                self.worksheet = None

            if self.excel_app:
                self.excel_app.Quit()
                logger.info("Đã thoát Excel Application")
                self.excel_app = None

        except Exception as e:
            logger.error(f"Lỗi khi đóng COM objects: {e}")
            self.workbook = None
            self.worksheet = None
            self.excel_app = None

