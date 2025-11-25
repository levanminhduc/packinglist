import tkinter as tk
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from excel_automation.utils import setup_logging
from config import settings
import logging

from ui.excel_viewer_window import ExcelViewerWindow

setup_logging(settings.LOG_FILE, getattr(logging, settings.LOG_LEVEL))
logger = logging.getLogger(__name__)


def main():
    logger.info("=== KHỞI ĐỘNG EXCEL VIEWER ===")
    
    try:
        root = tk.Tk()
        app = ExcelViewerWindow(root)
        root.mainloop()
        
    except Exception as e:
        logger.error(f"Lỗi khi khởi động ứng dụng: {e}", exc_info=True)
        print(f"❌ Lỗi: {e}")
    
    logger.info("=== ĐÓNG EXCEL VIEWER ===")


if __name__ == "__main__":
    main()

