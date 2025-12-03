import json
from typing import Dict, Any, Optional
import logging

from excel_automation.path_helper import get_config_path

logger = logging.getLogger(__name__)


class UIConfig:

    DEFAULT_CONFIG = {
        "window": {
            "width": 1200,
            "height": 800,
            "position_x": 100,
            "position_y": 50,
            "maximized": False
        },
        "table": {
            "column_width": 150,
            "row_height": 25,
            "header_height": 30,
            "font_family": "Arial",
            "font_size": 10,
            "show_grid": True
        },
        "theme": {
            "mode": "light",
            "primary_color": "#1f538d",
            "background_color": "#ffffff",
            "text_color": "#000000"
        },
        "recent_files": [],
        "last_opened_file": None
    }

    def __init__(self, config_file: str = "config/ui_config.json"):
        self.config_file = get_config_path(config_file)
        self.config: Dict[str, Any] = {}
        self._load_config()
    
    def _load_config(self) -> None:
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    self.config = self._merge_with_defaults(loaded_config)
                    logger.info(f"Đã tải cấu hình từ {self.config_file}")
            else:
                self.config = self.DEFAULT_CONFIG.copy()
                self._save_config()
                logger.info("Tạo cấu hình mặc định")
        except Exception as e:
            logger.error(f"Lỗi khi tải cấu hình: {e}")
            self.config = self.DEFAULT_CONFIG.copy()
    
    def _merge_with_defaults(self, loaded_config: Dict[str, Any]) -> Dict[str, Any]:
        merged = self.DEFAULT_CONFIG.copy()
        for key, value in loaded_config.items():
            if key in merged and isinstance(merged[key], dict) and isinstance(value, dict):
                merged[key].update(value)
            else:
                merged[key] = value
        return merged
    
    def _save_config(self) -> None:
        try:
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            logger.info(f"Đã lưu cấu hình vào {self.config_file}")
        except Exception as e:
            logger.error(f"Lỗi khi lưu cấu hình: {e}")
    
    def get(self, key: str, default: Any = None) -> Any:
        keys = key.split('.')
        value = self.config
        for k in keys:
            if isinstance(value, dict):
                value = value.get(k)
                if value is None:
                    return default
            else:
                return default
        return value
    
    def set(self, key: str, value: Any) -> None:
        keys = key.split('.')
        config = self.config
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        config[keys[-1]] = value
        self._save_config()
    
    def get_window_geometry(self) -> str:
        w = self.get('window.width', 1200)
        h = self.get('window.height', 800)
        x = self.get('window.position_x', 100)
        y = self.get('window.position_y', 50)
        return f"{w}x{h}+{x}+{y}"
    
    def set_window_geometry(self, geometry: str) -> None:
        try:
            parts = geometry.replace('+', 'x').split('x')
            if len(parts) >= 4:
                self.set('window.width', int(parts[0]))
                self.set('window.height', int(parts[1]))
                self.set('window.position_x', int(parts[2]))
                self.set('window.position_y', int(parts[3]))
        except Exception as e:
            logger.error(f"Lỗi khi lưu geometry: {e}")
    
    def add_recent_file(self, file_path: str) -> None:
        recent = self.get('recent_files', [])
        if file_path in recent:
            recent.remove(file_path)
        recent.insert(0, file_path)
        recent = recent[:10]
        self.set('recent_files', recent)
        self.set('last_opened_file', file_path)
    
    def get_recent_files(self) -> list:
        return self.get('recent_files', [])
    
    def get_table_config(self) -> Dict[str, Any]:
        return self.get('table', {})
    
    def update_table_config(self, **kwargs) -> None:
        for key, value in kwargs.items():
            self.set(f'table.{key}', value)
    
    def reset_to_defaults(self) -> None:
        self.config = self.DEFAULT_CONFIG.copy()
        self._save_config()
        logger.info("Đã reset cấu hình về mặc định")

