import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_automation.reader import ExcelReader
from excel_automation.validator import DataValidator
from excel_automation.utils import setup_logging, get_timestamp
from config.settings import Settings
import logging


def main():
    settings = Settings()
    
    log_file = settings.LOGS_DIR / f"validation_{get_timestamp()}.log"
    setup_logging(str(log_file), logging.INFO)
    logger = logging.getLogger(__name__)
    
    print("=" * 80)
    print("DATA VALIDATION ENGINE - DEMO")
    print("=" * 80)
    
    input_file = settings.get_input_path("sample_orders.xlsx")
    rules_file = Path("data/validation_rules/packing_list_rules.json")
    
    if not input_file.exists():
        print(f"❌ Không tìm thấy file input: {input_file}")
        print("💡 Chạy: python scripts/create_sample_data.py để tạo sample data")
        return
    
    if not rules_file.exists():
        print(f"❌ Không tìm thấy file rules: {rules_file}")
        return
    
    print(f"\n📂 Đọc dữ liệu từ: {input_file}")
    reader = ExcelReader(str(input_file))
    df = reader.read_with_pandas(sheet_name='Orders')
    print(f"✓ Đã đọc {len(df)} dòng dữ liệu")
    
    print(f"\n📋 Load validation rules từ: {rules_file}")
    validator = DataValidator.from_json(str(rules_file))
    print(f"✓ Đã load rules cho {len(validator.rules)} cột")
    
    print("\n🔍 Bắt đầu validation...")
    result = validator.validate_dataframe(df)
    
    print("\n" + "=" * 80)
    print("KẾT QUẢ VALIDATION")
    print("=" * 80)
    
    print(f"\n📊 Tổng quan:")
    print(f"  • Tổng số dòng: {result.total_rows}")
    print(f"  • Số dòng hợp lệ: {result.summary['valid_rows']}")
    print(f"  • Số lỗi: {result.error_count}")
    print(f"  • Trạng thái: {'✅ PASS' if result.is_valid else '❌ FAIL'}")
    
    if not result.is_valid:
        print(f"\n📋 Lỗi theo cột:")
        for column, count in result.summary['errors_by_column'].items():
            print(f"  • {column}: {count} lỗi")
        
        print(f"\n📝 Chi tiết lỗi (10 lỗi đầu tiên):")
        for i, error in enumerate(result.errors[:10], 1):
            print(f"\n  {i}. Dòng {error.row_index}, Cột '{error.column}'")
            print(f"     Giá trị: {error.value}")
            print(f"     Quy tắc: {error.rule}")
            print(f"     Lỗi: {error.message}")
        
        if len(result.errors) > 10:
            print(f"\n  ... và {len(result.errors) - 10} lỗi khác")
        
        error_report_path = settings.get_output_path(f"validation_errors_{get_timestamp()}.xlsx")
        print(f"\n💾 Tạo báo cáo lỗi...")
        validator.generate_error_report(result, str(error_report_path))
        print(f"✓ Đã tạo báo cáo lỗi tại: {error_report_path}")
        
        highlighted_path = settings.get_output_path(f"orders_highlighted_{get_timestamp()}.xlsx")
        print(f"\n🎨 Highlight lỗi trong file gốc...")
        validator.highlight_errors_in_excel(
            str(input_file),
            result,
            str(highlighted_path),
            sheet_name='Orders'
        )
        print(f"✓ Đã tạo file highlight tại: {highlighted_path}")
        
        print("\n" + "=" * 80)
        print("📁 OUTPUT FILES:")
        print(f"  1. Báo cáo lỗi: {error_report_path}")
        print(f"  2. File highlight: {highlighted_path}")
        print(f"  3. Log file: {log_file}")
        print("=" * 80)
    else:
        print("\n✅ Tất cả dữ liệu đều hợp lệ!")
        print("💡 Có thể tiếp tục xử lý dữ liệu này")
    
    print("\n✓ Hoàn thành!")


if __name__ == "__main__":
    main()

