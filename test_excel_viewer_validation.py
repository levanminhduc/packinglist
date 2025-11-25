import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

print("Testing Excel Viewer with Validation...")

try:
    from ui.excel_viewer_window import ExcelViewerWindow
    print("✅ ExcelViewerWindow imported successfully")
except Exception as e:
    print(f"❌ Failed to import ExcelViewerWindow: {e}")
    sys.exit(1)

try:
    from excel_automation import DataValidator
    print("✅ DataValidator imported successfully")
except Exception as e:
    print(f"❌ Failed to import DataValidator: {e}")
    sys.exit(1)

print("\n" + "="*60)
print("✅ ALL IMPORTS SUCCESSFUL!")
print("="*60)
print("\nBạn có thể chạy Excel Viewer với:")
print("  python excel_viewer.py")
print("\nTính năng mới:")
print("  1. Menu 'Validation' với các tùy chọn validation")
print("  2. Toolbar buttons: 'Load Rules' và 'Validate'")
print("  3. Keyboard shortcuts: Ctrl+L (Load Rules), Ctrl+V (Validate)")
print("  4. Highlight lỗi trực tiếp trong bảng (màu vàng)")
print("  5. Dialog hiển thị chi tiết lỗi")
print("  6. Export báo cáo lỗi ra Excel")
print("="*60)

