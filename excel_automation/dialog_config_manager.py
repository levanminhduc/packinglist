import json
from pathlib import Path
from typing import Dict, Any, Tuple, Optional
import logging

logger = logging.getLogger(__name__)


class DialogConfigManager:

    DEFAULT_CONFIG = {
        "main_window": {
            "width": 700,
            "height": 600,
            "x": None,
            "y": None
        },
        "dialogs": {
            "po_update": {
                "width": 400,
                "height": 300
            },
            "color_code_update": {
                "width": 450,
                "height": 400
            },
            "size_filter_config": {
                "width": 450,
                "height": 350
            }
        }
    }
    
    def __init__(self, config_file: str = "data/template_configs/dialog_config.json"):
        self.config_file = Path(config_file)
        self.config: Dict[str, Any] = {}
        self._load_config()
    
    def _load_config(self) -> None:
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                    logger.info(f"Đã tải cấu hình dialog từ {self.config_file}")
            else:
                self.config = self.DEFAULT_CONFIG.copy()
                self._save_config()
                logger.info("Tạo cấu hình dialog mặc định")
        except Exception as e:
            logger.error(f"Lỗi khi tải cấu hình dialog: {e}")
            self.config = self.DEFAULT_CONFIG.copy()
    
    def _save_config(self) -> None:
        try:
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            logger.info(f"Đã lưu cấu hình dialog vào {self.config_file}")
        except Exception as e:
            logger.error(f"Lỗi khi lưu cấu hình dialog: {e}")
    
    def get_dialog_size(self, dialog_name: str) -> Tuple[int, int]:
        try:
            dialog_config = self.config.get('dialogs', {}).get(dialog_name, {})
            width = dialog_config.get('width', 400)
            height = dialog_config.get('height', 300)
            return width, height
        except Exception as e:
            logger.error(f"Lỗi khi đọc kích thước dialog {dialog_name}: {e}")
            return 400, 300
    
    def save_dialog_size(self, dialog_name: str, width: int, height: int) -> None:
        try:
            if 'dialogs' not in self.config:
                self.config['dialogs'] = {}

            if dialog_name not in self.config['dialogs']:
                self.config['dialogs'][dialog_name] = {}

            self.config['dialogs'][dialog_name]['width'] = width
            self.config['dialogs'][dialog_name]['height'] = height

            self._save_config()
            logger.info(f"Đã lưu kích thước dialog {dialog_name}: {width}x{height}")
        except Exception as e:
            logger.error(f"Lỗi khi lưu kích thước dialog {dialog_name}: {e}")

    def get_main_window_geometry(self) -> Tuple[int, int, Optional[int], Optional[int]]:
        try:
            main_config = self.config.get('main_window', {})
            width = main_config.get('width', 700)
            height = main_config.get('height', 600)
            x = main_config.get('x')
            y = main_config.get('y')
            return width, height, x, y
        except Exception as e:
            logger.error(f"Lỗi khi đọc geometry cửa sổ chính: {e}")
            return 700, 600, None, None

    def save_main_window_geometry(self, width: int, height: int, x: int, y: int) -> None:
        try:
            if 'main_window' not in self.config:
                self.config['main_window'] = {}

            self.config['main_window']['width'] = width
            self.config['main_window']['height'] = height
            self.config['main_window']['x'] = x
            self.config['main_window']['y'] = y

            self._save_config()
            logger.info(f"Đã lưu geometry cửa sổ chính: {width}x{height}+{x}+{y}")
        except Exception as e:
            logger.error(f"Lỗi khi lưu geometry cửa sổ chính: {e}")

