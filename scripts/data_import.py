"""
Script import dữ liệu từ nhiều file Excel vào một file tổng hợp.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_automation import ExcelReader, ExcelWriter
from excel_automation.utils import setup_logging, list_excel_files, get_timestamp
from config import settings
import pandas as pd
import logging

setup_logging(settings.LOG_FILE, getattr(logging, settings.LOG_LEVEL))
logger = logging.getLogger(__name__)


def import_multiple_files():
    """Import dữ liệu từ nhiều file Excel."""
    try:
        logger.info("=== BẮT ĐẦU IMPORT DỮ LIỆU ===")
        
        input_dir = settings.DATA_INPUT_DIR
        excel_files = list_excel_files(str(input_dir))
        
        if not excel_files:
            logger.warning(f"Không tìm thấy file Excel nào trong {input_dir}")
            return
        
        logger.info(f"Tìm thấy {len(excel_files)} file Excel")
        
        all_data = []
        
        for file_path in excel_files:
            try:
                logger.info(f"Đang đọc: {Path(file_path).name}")
                reader = ExcelReader(file_path)
                df = reader.read_with_pandas()
                
                df['Source_File'] = Path(file_path).name
                
                all_data.append(df)
                logger.info(f"  ✓ Đọc thành công {len(df)} dòng")
                
            except Exception as e:
                logger.error(f"  ✗ Lỗi khi đọc {file_path}: {e}")
                continue
        
        if not all_data:
            logger.error("Không có dữ liệu nào được import")
            return
        
        combined_df = pd.concat(all_data, ignore_index=True)
        logger.info(f"Tổng cộng: {len(combined_df)} dòng từ {len(all_data)} file")
        
        timestamp = get_timestamp("%Y%m%d_%H%M%S")
        output_file = settings.get_output_path(f"combined_data_{timestamp}.xlsx")
        
        logger.info(f"Ghi dữ liệu tổng hợp ra: {output_file}")
        writer = ExcelWriter(str(output_file))
        writer.write_dataframe(combined_df, sheet_name='Combined Data')
        
        logger.info(f"✅ Hoàn thành! File output: {output_file}")
        logger.info("=== KẾT THÚC IMPORT DỮ LIỆU ===")
        
    except Exception as e:
        logger.error(f"❌ Lỗi khi import dữ liệu: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    import_multiple_files()

