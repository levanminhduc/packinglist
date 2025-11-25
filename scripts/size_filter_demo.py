import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_automation.size_filter import SizeFilterManager
from excel_automation.size_filter_config import SizeFilterConfig


def demo_basic_usage():
    print("=" * 60)
    print("DEMO 1: Sử dụng cơ bản SizeFilterManager")
    print("=" * 60)
    
    file_path = "data/sample.xlsx"
    
    if not Path(file_path).exists():
        print(f"⚠️  File không tồn tại: {file_path}")
        print("Vui lòng tạo file Excel mẫu với:")
        print("  - Cột F chứa size (044, 045, 046...)")
        print("  - Dữ liệu từ dòng 19 đến 59")
        return
    
    try:
        with SizeFilterManager(file_path) as manager:
            print(f"\n✓ Đã mở file: {file_path}")
            
            available_sizes = manager.scan_sizes()
            print(f"\n📊 Tìm thấy {len(available_sizes)} size khác nhau:")
            print(f"   {', '.join(available_sizes)}")
            
            size_rows = manager.get_size_row_mapping()
            print(f"\n📋 Chi tiết phân bố size:")
            for size, rows in sorted(size_rows.items()):
                print(f"   Size {size}: {len(rows)} dòng (dòng {min(rows)}-{max(rows)})")
            
            selected_sizes = available_sizes[:3] if len(available_sizes) >= 3 else available_sizes
            print(f"\n🔍 Áp dụng filter cho {len(selected_sizes)} size: {', '.join(selected_sizes)}")
            
            hidden_count = manager.apply_size_filter(selected_sizes)
            print(f"   ✓ Đã ẩn {hidden_count} dòng")
            
            output_path = "data/output/filtered_sample.xlsx"
            Path(output_path).parent.mkdir(parents=True, exist_ok=True)
            manager.save(output_path)
            print(f"\n💾 Đã lưu file: {output_path}")
            
    except Exception as e:
        print(f"\n❌ Lỗi: {e}")


def demo_custom_config():
    print("\n" + "=" * 60)
    print("DEMO 2: Sử dụng cấu hình tùy chỉnh")
    print("=" * 60)
    
    config = SizeFilterConfig()
    
    print(f"\n📝 Cấu hình hiện tại:")
    print(f"   Cột: {config.get_column()}")
    print(f"   Phạm vi: {config.get_start_row()} - {config.get_end_row()}")
    print(f"   Sheet: {config.get_sheet_name()}")
    
    try:
        print(f"\n🔧 Thử cập nhật cấu hình...")
        config.update_config("G", 20, 50, "Sheet2")
        print(f"   ✓ Đã cập nhật thành công")
        
        print(f"\n📝 Cấu hình mới:")
        print(f"   Cột: {config.get_column()}")
        print(f"   Phạm vi: {config.get_start_row()} - {config.get_end_row()}")
        print(f"   Sheet: {config.get_sheet_name()}")
        
        print(f"\n↩️  Reset về mặc định...")
        config.reset_to_defaults()
        print(f"   ✓ Đã reset")
        
    except Exception as e:
        print(f"\n❌ Lỗi: {e}")


def demo_validation():
    print("\n" + "=" * 60)
    print("DEMO 3: Validation cấu hình")
    print("=" * 60)
    
    config = SizeFilterConfig()
    
    test_cases = [
        ("Cấu hình hợp lệ", "F", 19, 59, None),
        ("start_row < 1", "F", 0, 59, None),
        ("start_row >= end_row", "F", 60, 59, None),
        ("end_row > max_row", "F", 19, 200, 100),
    ]
    
    for test_name, col, start, end, max_row in test_cases:
        print(f"\n🧪 Test: {test_name}")
        print(f"   Config: {col}[{start}:{end}], max_row={max_row}")
        
        try:
            config.update_config(col, start, end, "Sheet1")
            is_valid, msg = config.validate_config(max_row)
            
            if is_valid:
                print(f"   ✓ {msg}")
            else:
                print(f"   ⚠️  {msg}")
                
        except ValueError as e:
            print(f"   ❌ {e}")
        finally:
            config.reset_to_defaults()


def demo_reset_filter():
    print("\n" + "=" * 60)
    print("DEMO 4: Reset filter (hiện lại tất cả dòng)")
    print("=" * 60)
    
    file_path = "data/output/filtered_sample.xlsx"
    
    if not Path(file_path).exists():
        print(f"⚠️  File không tồn tại: {file_path}")
        print("Vui lòng chạy DEMO 1 trước")
        return
    
    try:
        with SizeFilterManager(file_path) as manager:
            print(f"\n✓ Đã mở file: {file_path}")
            
            print(f"\n🔄 Reset filter...")
            manager.reset_all_rows()
            print(f"   ✓ Đã hiện lại tất cả dòng")
            
            manager.save()
            print(f"\n💾 Đã lưu file")
            
    except Exception as e:
        print(f"\n❌ Lỗi: {e}")


def main():
    print("\n" + "=" * 60)
    print("SIZE FILTER DEMO - Tính năng lọc Size trong Excel")
    print("=" * 60)
    
    demo_basic_usage()
    demo_custom_config()
    demo_validation()
    demo_reset_filter()
    
    print("\n" + "=" * 60)
    print("✅ Hoàn thành tất cả demo")
    print("=" * 60)


if __name__ == "__main__":
    main()

