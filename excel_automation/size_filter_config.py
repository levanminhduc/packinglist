import json
from pathlib import Path
from typing import Dict, Any, Optional
import logging
import copy

logger = logging.getLogger(__name__)


class SizeFilterConfig:
    
    DEFAULT_CONFIG = {
        "size_filter_config": {
            "column": "F",
            "start_row": 19,
            "end_row": 59,
            "sheet_name": "Sheet1"
        }
    }
    
    def __init__(self, config_file: str = "data/template_configs/size_filter_config.json"):
        self.config_file = Path(config_file)
        self.config: Dict[str, Any] = {}
        self._load_config()
    
    def _load_config(self) -> None:
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    self.config = self._merge_with_defaults(loaded_config)
                    logger.info(f"Đã tải cấu hình size filter từ {self.config_file}")
            else:
                self.config = self.DEFAULT_CONFIG.copy()
                self._save_config()
                logger.info("Tạo cấu hình size filter mặc định")
        except Exception as e:
            logger.error(f"Lỗi khi tải cấu hình size filter: {e}")
            self.config = copy.deepcopy(self.DEFAULT_CONFIG)
    
    def _merge_with_defaults(self, loaded_config: Dict[str, Any]) -> Dict[str, Any]:
        merged = copy.deepcopy(self.DEFAULT_CONFIG)
        if 'size_filter_config' in loaded_config:
            merged['size_filter_config'].update(loaded_config['size_filter_config'])
        return merged
    
    def _save_config(self) -> None:
        try:
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            logger.info(f"Đã lưu cấu hình size filter vào {self.config_file}")
        except Exception as e:
            logger.error(f"Lỗi khi lưu cấu hình size filter: {e}")
            raise
    
    def get_column(self) -> str:
        return self.config['size_filter_config'].get('column', 'F')
    
    def get_start_row(self) -> int:
        return self.config['size_filter_config'].get('start_row', 19)
    
    def get_end_row(self) -> int:
        return self.config['size_filter_config'].get('end_row', 59)
    
    def get_sheet_name(self) -> str:
        return self.config['size_filter_config'].get('sheet_name', 'Sheet1')
    
    def set_column(self, column: str) -> None:
        if not column or len(column) > 3:
            raise ValueError("Cột phải là 1-3 ký tự (A-ZZZ)")
        self.config['size_filter_config']['column'] = column.upper()
        self._save_config()
    
    def set_start_row(self, start_row: int) -> None:
        if start_row < 1:
            raise ValueError("Dòng bắt đầu phải >= 1")
        if start_row >= self.get_end_row():
            raise ValueError("Dòng bắt đầu phải < dòng kết thúc")
        self.config['size_filter_config']['start_row'] = start_row
        self._save_config()
    
    def set_end_row(self, end_row: int) -> None:
        if end_row <= self.get_start_row():
            raise ValueError("Dòng kết thúc phải > dòng bắt đầu")
        self.config['size_filter_config']['end_row'] = end_row
        self._save_config()
    
    def set_sheet_name(self, sheet_name: str) -> None:
        if not sheet_name:
            raise ValueError("Tên sheet không được rỗng")
        self.config['size_filter_config']['sheet_name'] = sheet_name
        self._save_config()
    
    def update_config(self, column: str, start_row: int, end_row: int, sheet_name: str) -> None:
        if start_row < 1:
            raise ValueError("Dòng bắt đầu phải >= 1")
        if start_row >= end_row:
            raise ValueError("Dòng bắt đầu phải < dòng kết thúc")
        if not column or len(column) > 3:
            raise ValueError("Cột phải là 1-3 ký tự (A-ZZZ)")
        if not sheet_name:
            raise ValueError("Tên sheet không được rỗng")
        
        self.config['size_filter_config']['column'] = column.upper()
        self.config['size_filter_config']['start_row'] = start_row
        self.config['size_filter_config']['end_row'] = end_row
        self.config['size_filter_config']['sheet_name'] = sheet_name
        self._save_config()
        logger.info(f"Đã cập nhật config: {column} [{start_row}:{end_row}] sheet '{sheet_name}'")
    
    def validate_config(self, max_row: Optional[int] = None) -> tuple[bool, str]:
        try:
            start = self.get_start_row()
            end = self.get_end_row()
            
            if start < 1:
                return False, "Dòng bắt đầu phải >= 1"
            
            if start >= end:
                return False, "Dòng bắt đầu phải < dòng kết thúc"
            
            if max_row and end > max_row:
                return False, f"Dòng kết thúc ({end}) vượt quá số dòng thực tế ({max_row})"
            
            return True, "Config hợp lệ"
        except Exception as e:
            return False, f"Lỗi validate: {str(e)}"
    
    def reset_to_defaults(self) -> None:
        self.config = copy.deepcopy(self.DEFAULT_CONFIG)
        self._save_config()
        logger.info("Đã reset cấu hình size filter về mặc định")

