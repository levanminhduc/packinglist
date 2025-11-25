import json
from pathlib import Path
from typing import Dict, Any
import logging
import copy

logger = logging.getLogger(__name__)


class BoxListExportConfig:
    
    DEFAULT_CONFIG = {
        "box_list_export_config": {
            "box_start_row": 15,
            "box_end_row": 16,
            "size_column": "F",
            "size_data_start_row": 19,
            "size_data_end_row": 59,
            "combined_size_separator": "/",
            "enable_combined_detection": True,
            "sort_combined_sizes": True,
            "po_cell_row": 19,
            "po_cell_column": "A",
            "max_rows_per_column": 45,
            "header_rows": 2
        }
    }
    
    def __init__(self, config_file: str = "data/template_configs/box_list_export_config.json"):
        self.config_file = Path(config_file)
        self.config: Dict[str, Any] = {}
        self._load_config()
    
    def _load_config(self) -> None:
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    self.config = self._merge_with_defaults(loaded_config)
                    logger.info(f"Đã tải cấu hình box list export từ {self.config_file}")
            else:
                self.config = copy.deepcopy(self.DEFAULT_CONFIG)
                self._save_config()
                logger.info("Tạo cấu hình box list export mặc định")
        except Exception as e:
            logger.error(f"Lỗi khi tải cấu hình box list export: {e}")
            self.config = copy.deepcopy(self.DEFAULT_CONFIG)
    
    def _merge_with_defaults(self, loaded_config: Dict[str, Any]) -> Dict[str, Any]:
        merged = copy.deepcopy(self.DEFAULT_CONFIG)
        if 'box_list_export_config' in loaded_config:
            merged['box_list_export_config'].update(loaded_config['box_list_export_config'])
        return merged
    
    def _save_config(self) -> None:
        try:
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            logger.info(f"Đã lưu cấu hình box list export vào {self.config_file}")
        except Exception as e:
            logger.error(f"Lỗi khi lưu cấu hình box list export: {e}")
            raise
    
    def get_box_start_row(self) -> int:
        return self.config['box_list_export_config'].get('box_start_row', 15)
    
    def get_box_end_row(self) -> int:
        return self.config['box_list_export_config'].get('box_end_row', 16)
    
    def get_size_column(self) -> str:
        return self.config['box_list_export_config'].get('size_column', 'F')
    
    def get_size_data_start_row(self) -> int:
        return self.config['box_list_export_config'].get('size_data_start_row', 19)
    
    def get_size_data_end_row(self) -> int:
        return self.config['box_list_export_config'].get('size_data_end_row', 59)
    
    def get_combined_size_separator(self) -> str:
        return self.config['box_list_export_config'].get('combined_size_separator', '/')
    
    def is_combined_detection_enabled(self) -> bool:
        return self.config['box_list_export_config'].get('enable_combined_detection', True)
    
    def is_sort_combined_sizes_enabled(self) -> bool:
        return self.config['box_list_export_config'].get('sort_combined_sizes', True)

    def get_po_cell_row(self) -> int:
        return self.config['box_list_export_config'].get('po_cell_row', 19)

    def get_po_cell_column(self) -> str:
        return self.config['box_list_export_config'].get('po_cell_column', 'A')

    def get_max_rows_per_column(self) -> int:
        return self.config['box_list_export_config'].get('max_rows_per_column', 120)

    def get_header_rows(self) -> int:
        return self.config['box_list_export_config'].get('header_rows', 2)

