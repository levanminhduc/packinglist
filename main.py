"""
Main entry point cho Excel Automation.
Cung cấp menu tương tác để chọn các chức năng.
"""

import sys
from pathlib import Path
from excel_automation.utils import setup_logging
from config import settings
import logging

setup_logging(settings.LOG_FILE, getattr(logging, settings.LOG_LEVEL))
logger = logging.getLogger(__name__)


def print_menu():
    """In menu chính."""
    print("\n" + "="*60)
    print("           EXCEL AUTOMATION - MENU CHÍNH")
    print("="*60)
    print("\n📊 CÁC CHỨC NĂNG:")
    print("  1. Tạo báo cáo hàng ngày")
    print("  2. Import dữ liệu từ nhiều file")
    print("  3. Xử lý hàng loạt file Excel")
    print("  4. Demo đọc/ghi Excel đơn giản")
    print("  0. Thoát")
    print("\n" + "="*60)


def demo_read_write():
    """Demo đọc và ghi Excel đơn giản."""
    from excel_automation import ExcelReader, ExcelWriter, ExcelFormatter
    import pandas as pd
    
    print("\n📝 DEMO ĐỌC/GHI EXCEL")
    print("-" * 60)
    
    try:
        demo_data = {
            'Tên': ['Nguyễn Văn A', 'Trần Thị B', 'Lê Văn C'],
            'Tuổi': [25, 30, 28],
            'Lương': [10000000, 15000000, 12000000],
            'Phòng ban': ['IT', 'HR', 'IT']
        }
        
        df = pd.DataFrame(demo_data)
        print("\n✓ Tạo dữ liệu mẫu:")
        print(df)
        
        output_file = settings.get_output_path("demo_output.xlsx")
        
        print(f"\n✓ Ghi dữ liệu ra file: {output_file}")
        writer = ExcelWriter(str(output_file))
        writer.write_dataframe(df, sheet_name='Nhân viên')
        
        print("✓ Định dạng file Excel...")
        formatter = ExcelFormatter(str(output_file))
        formatter.format_header(bg_color="366092", font_color="FFFFFF")
        formatter.auto_adjust_column_width()
        formatter.add_borders()
        formatter.freeze_panes(row=1)
        
        print(f"\n✅ Hoàn thành! File đã được tạo tại: {output_file}")
        
        print("\n✓ Đọc lại file vừa tạo...")
        reader = ExcelReader(str(output_file))
        df_read = reader.read_with_pandas()
        print(df_read)
        
    except Exception as e:
        logger.error(f"Lỗi trong demo: {e}", exc_info=True)
        print(f"\n❌ Lỗi: {e}")


def run_daily_report():
    """Chạy script tạo báo cáo hàng ngày."""
    print("\n📊 CHẠY BÁO CÁO HÀNG NGÀY")
    print("-" * 60)
    
    try:
        from scripts.daily_report import create_daily_report
        create_daily_report()
    except Exception as e:
        logger.error(f"Lỗi khi chạy báo cáo: {e}", exc_info=True)
        print(f"\n❌ Lỗi: {e}")


def run_data_import():
    """Chạy script import dữ liệu."""
    print("\n📥 IMPORT DỮ LIỆU TỪ NHIỀU FILE")
    print("-" * 60)
    
    try:
        from scripts.data_import import import_multiple_files
        import_multiple_files()
    except Exception as e:
        logger.error(f"Lỗi khi import: {e}", exc_info=True)
        print(f"\n❌ Lỗi: {e}")


def run_batch_process():
    """Chạy script xử lý hàng loạt."""
    print("\n⚙️ XỬ LÝ HÀNG LOẠT FILE EXCEL")
    print("-" * 60)
    
    try:
        from scripts.batch_process import batch_process
        batch_process()
    except Exception as e:
        logger.error(f"Lỗi khi xử lý hàng loạt: {e}", exc_info=True)
        print(f"\n❌ Lỗi: {e}")


def main():
    """Hàm main chính."""
    logger.info("=== KHỞI ĐỘNG EXCEL AUTOMATION ===")
    
    print("\n🚀 Chào mừng đến với Excel Automation!")
    print(f"📁 Thư mục input: {settings.DATA_INPUT_DIR}")
    print(f"📁 Thư mục output: {settings.DATA_OUTPUT_DIR}")
    print(f"📝 Log file: {settings.LOG_FILE}")
    
    while True:
        print_menu()
        
        try:
            choice = input("\n👉 Chọn chức năng (0-4): ").strip()
            
            if choice == '0':
                print("\n👋 Tạm biệt!")
                logger.info("=== ĐÓNG EXCEL AUTOMATION ===")
                break
            
            elif choice == '1':
                run_daily_report()
            
            elif choice == '2':
                run_data_import()
            
            elif choice == '3':
                run_batch_process()
            
            elif choice == '4':
                demo_read_write()
            
            else:
                print("\n⚠️ Lựa chọn không hợp lệ! Vui lòng chọn từ 0-4.")
            
            input("\n⏎ Nhấn Enter để tiếp tục...")
            
        except KeyboardInterrupt:
            print("\n\n👋 Tạm biệt!")
            logger.info("=== ĐÓNG EXCEL AUTOMATION (KeyboardInterrupt) ===")
            break
        
        except Exception as e:
            logger.error(f"Lỗi không mong đợi: {e}", exc_info=True)
            print(f"\n❌ Lỗi: {e}")
            input("\n⏎ Nhấn Enter để tiếp tục...")


if __name__ == "__main__":
    main()

