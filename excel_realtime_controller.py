import tkinter as tk
import logging
import sys
from pathlib import Path

from ui.excel_realtime_controller import ExcelRealtimeController

def get_base_path() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent

base_path = get_base_path()
logs_dir = base_path / 'logs'
logs_dir.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(logs_dir / 'realtime_controller.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)


def main():
    logger.info("Khởi động Excel Real-Time Controller")
    
    root = tk.Tk()
    app = ExcelRealtimeController(root)
    root.mainloop()
    
    logger.info("Đã đóng Excel Real-Time Controller")


if __name__ == "__main__":
    main()

