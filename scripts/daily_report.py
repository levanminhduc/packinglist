"""
Script tạo báo cáo hàng ngày từ dữ liệu Excel.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_automation import ExcelReader, ExcelWriter, ExcelProcessor, ExcelFormatter
from excel_automation.utils import setup_logging, get_timestamp, create_backup
from config import settings
import logging

setup_logging(settings.LOG_FILE, getattr(logging, settings.LOG_LEVEL))
logger = logging.getLogger(__name__)


def create_daily_report():
    """Tạo báo cáo hàng ngày."""
    try:
        logger.info("=== BẮT ĐẦU TẠO BÁO CÁO HÀNG NGÀY ===")
        
        input_file = settings.get_input_path("sales_data.xlsx")
        
        if not input_file.exists():
            logger.error(f"File input không tồn tại: {input_file}")
            return
        
        logger.info(f"Đọc dữ liệu từ: {input_file}")
        reader = ExcelReader(str(input_file))
        df = reader.read_with_pandas()
        
        logger.info(f"Đã đọc {len(df)} dòng dữ liệu")
        
        processor = ExcelProcessor()
        
        df_clean = processor.clean_data(df, drop_duplicates=True, drop_na=False)
        logger.info(f"Sau khi làm sạch: {len(df_clean)} dòng")
        
        df_summary = processor.aggregate_data(
            df_clean,
            group_by=['Category'],
            agg_dict={'Amount': 'sum', 'Quantity': 'sum'}
        )
        
        df_sorted = processor.sort_data(df_summary, by=['Amount'], ascending=False)
        
        timestamp = get_timestamp("%Y%m%d")
        output_file = settings.get_output_path(f"daily_report_{timestamp}.xlsx")
        
        logger.info(f"Ghi báo cáo ra: {output_file}")
        writer = ExcelWriter(str(output_file))
        
        writer.write_multiple_sheets({
            'Summary': df_sorted,
            'Raw Data': df_clean
        })
        
        formatter = ExcelFormatter(str(output_file))
        formatter.format_header(sheet_name='Summary')
        formatter.auto_adjust_column_width(sheet_name='Summary')
        formatter.add_borders(sheet_name='Summary')
        formatter.freeze_panes(sheet_name='Summary', row=1)
        
        logger.info(f"✅ Hoàn thành! File báo cáo: {output_file}")
        logger.info("=== KẾT THÚC TẠO BÁO CÁO ===")
        
    except Exception as e:
        logger.error(f"❌ Lỗi khi tạo báo cáo: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    create_daily_report()

