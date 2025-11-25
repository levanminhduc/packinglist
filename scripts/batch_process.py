"""
Script xử lý hàng loạt file Excel với các thao tác tùy chỉnh.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_automation import ExcelReader, ExcelWriter, ExcelProcessor, ExcelFormatter
from excel_automation.utils import setup_logging, list_excel_files, create_backup, get_timestamp
from config import settings
import logging

setup_logging(settings.LOG_FILE, getattr(logging, settings.LOG_LEVEL))
logger = logging.getLogger(__name__)


def process_single_file(file_path: str) -> bool:
    """
    Xử lý một file Excel.
    
    Args:
        file_path: Đường dẫn file cần xử lý
        
    Returns:
        True nếu thành công
    """
    try:
        file_name = Path(file_path).name
        logger.info(f"Xử lý file: {file_name}")
        
        if settings.AUTO_BACKUP:
            backup_path = create_backup(file_path, str(settings.DATA_BACKUP_DIR))
            logger.info(f"  Đã backup: {Path(backup_path).name}")
        
        reader = ExcelReader(file_path)
        df = reader.read_with_pandas()
        logger.info(f"  Đọc {len(df)} dòng")
        
        processor = ExcelProcessor()
        
        df_clean = processor.clean_data(df, drop_duplicates=True, fill_na=0)
        
        df_sorted = processor.sort_data(df_clean, by=[df_clean.columns[0]], ascending=True)
        
        output_file = settings.get_output_path(f"processed_{file_name}")
        
        writer = ExcelWriter(str(output_file))
        writer.write_dataframe(df_sorted, sheet_name='Processed Data')
        
        formatter = ExcelFormatter(str(output_file))
        formatter.format_header()
        formatter.auto_adjust_column_width()
        formatter.add_borders()
        formatter.freeze_panes(row=1)
        
        logger.info(f"  ✓ Hoàn thành: {output_file.name}")
        return True
        
    except Exception as e:
        logger.error(f"  ✗ Lỗi khi xử lý {file_path}: {e}")
        return False


def batch_process():
    """Xử lý hàng loạt tất cả file Excel trong thư mục input."""
    try:
        logger.info("=== BẮT ĐẦU XỬ LÝ HÀNG LOẠT ===")
        
        input_dir = settings.DATA_INPUT_DIR
        excel_files = list_excel_files(str(input_dir))
        
        if not excel_files:
            logger.warning(f"Không tìm thấy file Excel nào trong {input_dir}")
            return
        
        logger.info(f"Tìm thấy {len(excel_files)} file cần xử lý")
        
        success_count = 0
        fail_count = 0
        
        for file_path in excel_files:
            if process_single_file(file_path):
                success_count += 1
            else:
                fail_count += 1
        
        logger.info("=== KẾT QUẢ XỬ LÝ ===")
        logger.info(f"✅ Thành công: {success_count} file")
        logger.info(f"❌ Thất bại: {fail_count} file")
        logger.info(f"📊 Tổng cộng: {len(excel_files)} file")
        logger.info("=== KẾT THÚC XỬ LÝ HÀNG LOẠT ===")
        
    except Exception as e:
        logger.error(f"❌ Lỗi trong quá trình xử lý hàng loạt: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    batch_process()

