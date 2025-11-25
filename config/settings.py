"""
Cấu hình settings cho ứng dụng.
Sử dụng environment variables với fallback defaults.
"""

import os
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv

load_dotenv()

PROJECT_ROOT = Path(__file__).parent.parent
DATA_DIR = PROJECT_ROOT / "data"
LOGS_DIR = PROJECT_ROOT / "logs"


class Settings:
    """Class quản lý settings của ứng dụng."""
    
    APP_NAME: str = "Excel Automation"
    APP_VERSION: str = "1.0.0"
    DEBUG: bool = os.getenv("DEBUG", "False").lower() == "true"
    
    DATA_INPUT_DIR: Path = Path(os.getenv("DATA_INPUT_DIR", DATA_DIR / "input"))
    DATA_OUTPUT_DIR: Path = Path(os.getenv("DATA_OUTPUT_DIR", DATA_DIR / "output"))
    DATA_TEMPLATES_DIR: Path = Path(os.getenv("DATA_TEMPLATES_DIR", DATA_DIR / "templates"))
    DATA_BACKUP_DIR: Path = Path(os.getenv("DATA_BACKUP_DIR", DATA_DIR / "backup"))
    
    LOGS_DIR: Path = Path(os.getenv("LOGS_DIR", LOGS_DIR))
    LOG_FILE: str = os.getenv("LOG_FILE", str(LOGS_DIR / "app.log"))
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "INFO")
    
    MAX_ROWS: int = int(os.getenv("MAX_ROWS", "1000000"))
    DEFAULT_SHEET_NAME: str = os.getenv("DEFAULT_SHEET_NAME", "Sheet1")
    
    AUTO_BACKUP: bool = os.getenv("AUTO_BACKUP", "True").lower() == "true"
    BACKUP_KEEP_DAYS: int = int(os.getenv("BACKUP_KEEP_DAYS", "30"))
    
    EXCEL_ENGINE: str = os.getenv("EXCEL_ENGINE", "openpyxl")
    
    DATE_FORMAT: str = os.getenv("DATE_FORMAT", "%Y-%m-%d")
    DATETIME_FORMAT: str = os.getenv("DATETIME_FORMAT", "%Y-%m-%d %H:%M:%S")
    
    @classmethod
    def validate(cls):
        """Validate và tạo các thư mục cần thiết."""
        cls.DATA_INPUT_DIR.mkdir(parents=True, exist_ok=True)
        cls.DATA_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        cls.DATA_TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
        cls.DATA_BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        cls.LOGS_DIR.mkdir(parents=True, exist_ok=True)
    
    @classmethod
    def get_input_path(cls, filename: str) -> Path:
        """Lấy đường dẫn file trong thư mục input."""
        return cls.DATA_INPUT_DIR / filename
    
    @classmethod
    def get_output_path(cls, filename: str) -> Path:
        """Lấy đường dẫn file trong thư mục output."""
        return cls.DATA_OUTPUT_DIR / filename
    
    @classmethod
    def get_template_path(cls, filename: str) -> Path:
        """Lấy đường dẫn file trong thư mục templates."""
        return cls.DATA_TEMPLATES_DIR / filename
    
    @classmethod
    def get_backup_path(cls, filename: str) -> Path:
        """Lấy đường dẫn file trong thư mục backup."""
        return cls.DATA_BACKUP_DIR / filename


settings = Settings()
settings.validate()

