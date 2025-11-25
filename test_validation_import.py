import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

print("Testing imports...")

try:
    from excel_automation import DataValidator
    print("✅ DataValidator imported")
except Exception as e:
    print(f"❌ DataValidator import failed: {e}")
    sys.exit(1)

try:
    from excel_automation import ValidationResult
    print("✅ ValidationResult imported")
except Exception as e:
    print(f"❌ ValidationResult import failed: {e}")
    sys.exit(1)

try:
    from excel_automation import ValidationError
    print("✅ ValidationError imported")
except Exception as e:
    print(f"❌ ValidationError import failed: {e}")
    sys.exit(1)

try:
    from excel_automation import RequiredRule, TypeRule, RangeRule
    print("✅ Validation rules imported")
except Exception as e:
    print(f"❌ Validation rules import failed: {e}")
    sys.exit(1)

print("\n" + "="*50)
print("✅ ALL IMPORTS SUCCESSFUL!")
print("="*50)

