from typing import List, Optional, Set
from pathlib import Path
import logging
import win32com.client
from win32com.client import CDispatch

from excel_automation.size_filter_config import SizeFilterConfig
from excel_automation.utils import get_size_sort_key

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
    
    def open_excel_file(self, file_path: str) -> None:
        file_path_obj = Path(file_path)
        if not file_path_obj.exists():
            raise FileNotFoundError(f"File không tồn tại: {file_path}")
        
        try:
            if self.excel_app is None:
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

            self.workbook = self.excel_app.Workbooks.Open(str(file_path_obj.absolute()))
            self.current_file = file_path
            
            sheet_name = self.config.get_sheet_name()
            try:
                self.worksheet = self.workbook.Sheets(sheet_name)
                self.current_sheet = sheet_name
            except Exception:
                self.worksheet = self.workbook.Sheets(1)
                self.current_sheet = self.worksheet.Name
                logger.warning(f"Sheet '{sheet_name}' không tồn tại, sử dụng sheet đầu tiên: {self.current_sheet}")
            
            self.worksheet.Activate()
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
    
    def scan_sizes(self, column: Optional[str] = None, start_row: Optional[int] = None,
                   end_row: Optional[int] = None) -> List[str]:
        if self.worksheet is None:
            raise RuntimeError("Chưa chọn worksheet nào")
        
        column = column or self.config.get_column()
        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.config.get_end_row()
        
        try:
            sizes: Set[str] = set()
            
            for row in range(start_row, end_row + 1):
                cell_value = self.worksheet.Cells(row, self._column_letter_to_number(column)).Value
                
                if cell_value is not None:
                    size_str = str(cell_value).strip()
                    
                    if size_str.isdigit():
                        size_str = size_str.zfill(3)
                    
                    if size_str:
                        sizes.add(size_str)

            sorted_sizes = sorted(sizes, key=get_size_sort_key)
            logger.info(f"Quét được {len(sorted_sizes)} size khác nhau trong {column}[{start_row}:{end_row}]")
            return sorted_sizes
            
        except Exception as e:
            logger.error(f"Lỗi khi quét sizes: {e}")
            raise RuntimeError(f"Không thể quét sizes: {str(e)}")
    
    def hide_rows_realtime(self, selected_sizes: List[str], column: Optional[str] = None,
                          start_row: Optional[int] = None, end_row: Optional[int] = None) -> int:
        if self.worksheet is None:
            raise RuntimeError("Chưa chọn worksheet nào")
        
        column = column or self.config.get_column()
        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.config.get_end_row()
        
        try:
            if self.excel_app:
                self.excel_app.ScreenUpdating = False
            
            selected_set = set(selected_sizes)
            hidden_count = 0
            col_num = self._column_letter_to_number(column)
            
            for row in range(start_row, end_row + 1):
                cell_value = self.worksheet.Cells(row, col_num).Value
                
                if cell_value is not None:
                    size_str = str(cell_value).strip()
                    if size_str.isdigit():
                        size_str = size_str.zfill(3)
                    
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
        end_row = end_row or self.config.get_end_row()
        
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
    
    def _column_letter_to_number(self, column: str) -> int:
        column = column.upper()
        result = 0
        for char in column:
            result = result * 26 + (ord(char) - ord('A') + 1)
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

