# Data Validation Engine

Module validation dữ liệu mạnh mẽ cho Excel với hỗ trợ nhiều loại quy tắc validation.

## Tính năng

- ✅ Validation rules đa dạng (Required, Type, Range, Regex, Length, Date, Unique, InSet, Custom)
- 📊 Báo cáo lỗi chi tiết với thông tin dòng, cột, giá trị, quy tắc vi phạm
- 🎨 Highlight lỗi trực tiếp trong file Excel với màu sắc và comments
- 📋 Load rules từ JSON config file
- 🔍 Validation kết quả với summary thống kê

## Cài đặt

Các dependencies đã được cài đặt sẵn trong project.

## Sử dụng nhanh

### 1. Tạo Validator từ JSON config

```python
from excel_automation import DataValidator

validator = DataValidator.from_json('data/validation_rules/packing_list_rules.json')
```

### 2. Validate DataFrame

```python
from excel_automation import ExcelReader

reader = ExcelReader('data/input/orders.xlsx')
df = reader.read_with_pandas()

result = validator.validate_dataframe(df)

if result.is_valid:
    print("✅ Dữ liệu hợp lệ!")
else:
    print(f"❌ Có {result.error_count} lỗi")
```

### 3. Tạo báo cáo lỗi

```python
if not result.is_valid:
    validator.generate_error_report(result, 'output/errors.xlsx')
    validator.highlight_errors_in_excel(
        'input/orders.xlsx',
        result,
        'output/orders_highlighted.xlsx'
    )
```

## Validation Rules

### RequiredRule

Kiểm tra trường bắt buộc không được để trống.

```python
from excel_automation import RequiredRule

rule = RequiredRule('PO', error_message='Số PO là bắt buộc')
```

### TypeRule

Kiểm tra kiểu dữ liệu (int, float, str).

```python
from excel_automation import TypeRule

rule = TypeRule('Quantity', expected_type=int, error_message='Quantity phải là số nguyên')
```

### RangeRule

Kiểm tra giá trị nằm trong khoảng min-max.

```python
from excel_automation import RangeRule

rule = RangeRule('Quantity', min_value=1, max_value=100000)
```

### RegexRule

Kiểm tra giá trị khớp với regex pattern.

```python
from excel_automation import RegexRule

rule = RegexRule('PO', pattern=r'^PO\d{7}$', error_message='PO phải có định dạng PO + 7 chữ số')
```

### LengthRule

Kiểm tra độ dài chuỗi.

```python
from excel_automation import LengthRule

rule = LengthRule('Style', min_length=3, max_length=20)
```

### DateRule

Kiểm tra định dạng ngày tháng.

```python
from excel_automation import DateRule

rule = DateRule('ShipDate', date_format='%Y-%m-%d')
```

### UniqueRule

Kiểm tra giá trị không bị trùng lặp.

```python
from excel_automation import UniqueRule

rule = UniqueRule('PO', error_message='Số PO bị trùng lặp')
```

### InSetRule

Kiểm tra giá trị nằm trong danh sách cho phép.

```python
from excel_automation import InSetRule

rule = InSetRule('Size', allowed_values=['XS', 'S', 'M', 'L', 'XL', 'XXL'], case_sensitive=False)
```

### CustomRule

Tạo rule tùy chỉnh với function.

```python
from excel_automation import CustomRule

def validate_po_prefix(value, row_data):
    return str(value).startswith('PO')

rule = CustomRule('PO', validation_func=validate_po_prefix)
```

## JSON Config Format

File `data/validation_rules/packing_list_rules.json`:

```json
{
  "PO": [
    {
      "type": "required",
      "error_message": "Số PO là bắt buộc"
    },
    {
      "type": "regex",
      "params": {
        "pattern": "^PO\\d{7}$"
      },
      "error_message": "Số PO phải có định dạng PO + 7 chữ số"
    }
  ],
  "Quantity": [
    {
      "type": "required",
      "error_message": "Quantity là bắt buộc"
    },
    {
      "type": "type",
      "params": {
        "expected_type": "int"
      },
      "error_message": "Quantity phải là số nguyên"
    },
    {
      "type": "range",
      "params": {
        "min_value": 1,
        "max_value": 100000
      },
      "error_message": "Quantity phải nằm trong khoảng từ 1 đến 100,000"
    }
  ]
}
```

## ValidationResult

Object chứa kết quả validation:

```python
result = validator.validate_dataframe(df)

print(result.is_valid)        # True/False
print(result.total_rows)      # Tổng số dòng
print(result.error_count)     # Số lỗi
print(result.errors)          # List[ValidationError]
print(result.summary)         # Dict với thống kê chi tiết
```

## ValidationError

Object chứa thông tin lỗi:

```python
for error in result.errors:
    print(f"Dòng {error.row_index}")
    print(f"Cột {error.column}")
    print(f"Giá trị {error.value}")
    print(f"Quy tắc {error.rule}")
    print(f"Lỗi {error.message}")
```

## Demo Script

Chạy demo validation:

```bash
python scripts/validate_data.py
```

Script sẽ:
1. Đọc file `data/input/sample_orders.xlsx`
2. Load rules từ `data/validation_rules/packing_list_rules.json`
3. Validate dữ liệu
4. Tạo báo cáo lỗi nếu có
5. Highlight lỗi trong file Excel

## Tạo Sample Data

Tạo sample data để test:

```bash
python scripts/create_sample_data.py
```

## Best Practices

1. **Định nghĩa rules rõ ràng**: Sử dụng error_message cụ thể cho từng rule
2. **Sắp xếp rules theo thứ tự**: Required → Type → Range/Length/Regex
3. **Sử dụng JSON config**: Dễ maintain và update rules
4. **Test với nhiều trường hợp**: Valid và invalid data
5. **Log validation results**: Để tracking và debugging

## Troubleshooting

### Lỗi "Column not found"

Đảm bảo tên cột trong rules JSON khớp với tên cột trong DataFrame.

### Lỗi "Invalid regex pattern"

Kiểm tra regex pattern có đúng syntax không. Nhớ escape các ký tự đặc biệt.

### Performance với file lớn

Với file > 100k rows, consider:
- Validate theo batch
- Sử dụng multiprocessing
- Tắt highlight errors (chỉ tạo báo cáo)

