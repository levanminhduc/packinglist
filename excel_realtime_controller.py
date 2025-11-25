import tkinter as tk
import logging
from pathlib import Path

from ui.excel_realtime_controller import ExcelRealtimeController

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/realtime_controller.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)


def main():
    logger.info("Khởi động Excel Real-Time Controller")
    
    logs_dir = Path('logs')
    logs_dir.mkdir(exist_ok=True)
    
    root = tk.Tk()
    app = ExcelRealtimeController(root)
    root.mainloop()
    
    logger.info("Đã đóng Excel Real-Time Controller")


if __name__ == "__main__":
    main()

