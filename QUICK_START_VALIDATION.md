# Quick Start - Data Validation Engine

## Chạy Demo ngay lập tức

```bash
# Bước 1: Tạo sample data (nếu chưa có)
python scripts/create_sample_data.py

# Bước 2: Chạy validation
python scripts/validate_data.py
```

## Kết quả mong đợi

```
================================================================================
DATA VALIDATION ENGINE - DEMO
================================================================================

📂 Đọc dữ liệu từ: D:\DuAnMoi\PythonExcel\data\input\sample_orders.xlsx
✓ Đã đọc 25 dòng dữ liệu

📋 Load validation rules từ: data\validation_rules\packing_list_rules.json
✓ Đã load rules cho 8 cột

🔍 Bắt đầu validation...

================================================================================
KẾT QUẢ VALIDATION
================================================================================

📊 Tổng quan:
  • Tổng số dòng: 25
  • Số dòng hợp lệ: 15
  • Số lỗi: 11
  • Trạng thái: ❌ FAIL

📋 Lỗi theo cột:
  • PO: 1 lỗi
  • Style: 1 lỗi
  • Color: 1 lỗi
  • Size: 1 lỗi
  • Quantity: 4 lỗi
  • ShipDate: 1 lỗi
  • Buyer: 1 lỗi
  • Carton: 1 lỗi

💾 Tạo báo cáo lỗi...
✓ Đã tạo báo cáo lỗi tại: data\output\validation_errors_*.xlsx

🎨 Highlight lỗi trong file gốc...
✓ Đã tạo file highlight tại: data\output\orders_highlighted_*.xlsx

================================================================================
📁 OUTPUT FILES:
  1. Báo cáo lỗi: data\output\validation_errors_*.xlsx
  2. File highlight: data\output\orders_highlighted_*.xlsx
  3. Log file: logs\validation_*.log
================================================================================

✓ Hoàn thành!
```

## Sử dụng trong code của bạn

### 1. Validate file Excel đơn giản

```python
from excel_automation import DataValidator, ExcelReader

# Đọc file
reader = ExcelReader('data/input/your_file.xlsx')
df = reader.read_with_pandas()

# Load validator từ JSON config
validator = DataValidator.from_json('data/validation_rules/packing_list_rules.json')

# Validate
result = validator.validate_dataframe(df)

# Kiểm tra kết quả
if result.is_valid:
    print("✅ Dữ liệu hợp lệ!")
    # Tiếp tục xử lý...
else:
    print(f"❌ Có {result.error_count} lỗi")
    # Tạo báo cáo
    validator.generate_error_report(result, 'output/errors.xlsx')
```

### 2. Tạo validator với code (không dùng JSON)

```python
from excel_automation import (
    DataValidator, 
    RequiredRule, 
    RegexRule, 
    RangeRule,
    InSetRule
)

validator = DataValidator()

# Thêm rules cho cột PO
validator.add_rule('PO', RequiredRule('PO', 'Số PO là bắt buộc'))
validator.add_rule('PO', RegexRule('PO', r'^PO\d{7}$', 'PO phải có format PO + 7 số'))

# Thêm rules cho cột Size
validator.add_rule('Size', RequiredRule('Size'))
validator.add_rule('Size', InSetRule('Size', ['XS', 'S', 'M', 'L', 'XL', 'XXL'], case_sensitive=False))

# Thêm rules cho cột Quantity
validator.add_rule('Quantity', RequiredRule('Quantity'))
validator.add_rule('Quantity', RangeRule('Quantity', min_value=1, max_value=100000))

# Validate
result = validator.validate_dataframe(df)
```

### 3. Xử lý kết quả validation chi tiết

```python
result = validator.validate_dataframe(df)

# In summary
print(f"Valid: {result.is_valid}")
print(f"Total rows: {result.total_rows}")
print(f"Errors: {result.error_count}")
print(f"Valid rows: {result.summary['valid_rows']}")

# In lỗi theo cột
for column, count in result.summary['errors_by_column'].items():
    print(f"{column}: {count} lỗi")

# In chi tiết từng lỗi
for error in result.errors:
    print(f"Dòng {error.row_index}: {error.column} = {error.value}")
    print(f"  Lỗi: {error.message}")
```

### 4. Tạo báo cáo và highlight lỗi

```python
if not result.is_valid:
    # Tạo báo cáo Excel
    validator.generate_error_report(
        result, 
        'output/error_report.xlsx'
    )
    
    # Highlight lỗi trong file gốc
    validator.highlight_errors_in_excel(
        'input/original.xlsx',
        result,
        'output/highlighted.xlsx',
        sheet_name='Orders'
    )
```

## Tùy chỉnh Validation Rules

### Tạo file JSON rules mới

Tạo file `my_rules.json`:

```json
{
  "ColumnName": [
    {
      "type": "required",
      "error_message": "Trường này là bắt buộc"
    },
    {
      "type": "type",
      "params": {
        "expected_type": "int"
      },
      "error_message": "Phải là số nguyên"
    },
    {
      "type": "range",
      "params": {
        "min_value": 0,
        "max_value": 1000
      },
      "error_message": "Giá trị phải từ 0 đến 1000"
    }
  ]
}
```

Sử dụng:

```python
validator = DataValidator.from_json('my_rules.json')
```

## Các loại Rules có sẵn

| Rule Type | Mô tả | Params |
|-----------|-------|--------|
| `required` | Trường bắt buộc | Không |
| `type` | Kiểm tra kiểu dữ liệu | `expected_type`: "int", "float", "str" |
| `range` | Giá trị trong khoảng | `min_value`, `max_value` |
| `regex` | Khớp với pattern | `pattern`: regex string |
| `length` | Độ dài chuỗi | `min_length`, `max_length` |
| `date` | Định dạng ngày | `date_format`: "%Y-%m-%d" |
| `unique` | Không trùng lặp | Không |
| `in_set` | Trong danh sách | `allowed_values`: list, `case_sensitive`: bool |

## Tips & Tricks

### 1. Validate nhiều files

```python
import glob

validator = DataValidator.from_json('rules.json')

for file_path in glob.glob('input/*.xlsx'):
    reader = ExcelReader(file_path)
    df = reader.read_with_pandas()
    result = validator.validate_dataframe(df)
    
    if not result.is_valid:
        print(f"❌ {file_path}: {result.error_count} lỗi")
    else:
        print(f"✅ {file_path}: OK")
```

### 2. Chỉ validate một số cột

```python
# Chỉ load rules cho cột cần thiết
validator = DataValidator()
validator.add_rule('PO', RequiredRule('PO'))
validator.add_rule('Quantity', RangeRule('Quantity', 1, 100000))
```

### 3. Custom error messages

```python
rule = RequiredRule(
    'PO', 
    error_message='⚠️ Vui lòng nhập số PO!'
)
```

### 4. Validate trước khi import vào database

```python
result = validator.validate_dataframe(df)

if result.is_valid:
    # Import vào database
    df.to_sql('orders', engine, if_exists='append')
else:
    # Gửi email báo lỗi
    send_error_email(result)
```

## Troubleshooting

### Lỗi: "Column not found"

**Nguyên nhân**: Tên cột trong rules không khớp với DataFrame

**Giải pháp**:
```python
# Kiểm tra tên cột
print(df.columns.tolist())

# Đảm bảo tên cột trong JSON khớp chính xác
```

### Lỗi: "Invalid regex pattern"

**Nguyên nhân**: Regex pattern không đúng syntax

**Giải pháp**:
```python
# Test regex trước
import re
pattern = r'^PO\d{7}$'
re.compile(pattern)  # Sẽ raise error nếu pattern sai
```

### Performance chậm với file lớn

**Giải pháp**:
```python
# Validate theo batch
chunk_size = 10000
for chunk in pd.read_excel('large_file.xlsx', chunksize=chunk_size):
    result = validator.validate_dataframe(chunk)
    # Process result...
```

## Xem thêm

- 📖 [Tài liệu đầy đủ](docs/VALIDATION_ENGINE.md)
- 📋 [Implementation Summary](IMPLEMENTATION_SUMMARY.md)
- 💻 [Demo Script](scripts/validate_data.py)

