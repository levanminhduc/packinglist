# Hướng Dẫn Sử Dụng Tính Năng Lọc Size

## Tổng Quan

Tính năng lọc size cho phép bạn:
- Quét và thu thập danh sách các size trong cột Excel
- Chọn size nào muốn hiển thị
- Ẩn các dòng có size không được chọn
- Cấu hình phạm vi quét linh hoạt

## Cấu Hình Mặc Định

File config: `data/template_configs/size_filter_config.json`

```json
{
  "size_filter_config": {
    "column": "F",
    "start_row": 19,
    "end_row": 59,
    "sheet_name": "Sheet1"
  }
}
```

## Sử Dụng Qua Giao Diện

### 1. Mở File Excel
- Menu: **File** → **Mở File Excel...**
- Hoặc nhấn `Ctrl+O`

### 2. Lọc Size
- Menu: **Lọc Size** → **Lọc theo Size...**
- Hoặc nhấn `Ctrl+F`
- Chọn các size muốn hiển thị
- Nhấn **Áp dụng**

### 3. Cấu Hình Phạm Vi
- Menu: **Lọc Size** → **Cấu hình Lọc Size...**
- Chỉnh sửa:
  - Tên sheet
  - Cột chứa size
  - Dòng bắt đầu
  - Dòng kết thúc
- Nhấn **Lưu**

### 4. Reset Lọc
- Menu: **Lọc Size** → **Reset Lọc Size**
- Hiện lại tất cả dòng đã bị ẩn

## Sử Dụng Qua Code

### Ví Dụ 1: Lọc Size Cơ Bản

```python
from excel_automation import SizeFilterManager

with SizeFilterManager("file.xlsx") as manager:
    # Quét sizes
    sizes = manager.scan_sizes()
    print(f"Tìm thấy: {sizes}")
    
    # Chọn size muốn hiển thị
    selected = ["044", "045", "046"]
    
    # Áp dụng filter
    hidden_count = manager.apply_size_filter(selected)
    print(f"Đã ẩn {hidden_count} dòng")
    
    # Lưu file
    manager.save()
```

### Ví Dụ 2: Cấu Hình Tùy Chỉnh

```python
from excel_automation import SizeFilterManager, SizeFilterConfig

# Tạo config tùy chỉnh
config = SizeFilterConfig()
config.update_config(
    column="G",
    start_row=20,
    end_row=50,
    sheet_name="Sheet2"
)

# Sử dụng config
with SizeFilterManager("file.xlsx", config) as manager:
    sizes = manager.scan_sizes()
    manager.apply_size_filter(sizes[:5])
    manager.save()
```

### Ví Dụ 3: Reset Filter

```python
from excel_automation import SizeFilterManager

with SizeFilterManager("file.xlsx") as manager:
    # Hiện lại tất cả dòng
    manager.reset_all_rows()
    manager.save()
```

## Validation

### Quy Tắc Validation

1. **Dòng bắt đầu**: Phải >= 1
2. **Dòng kết thúc**: Phải > dòng bắt đầu
3. **Phạm vi ẩn**: CHỈ ẩn dòng trong khoảng `start_row` đến `end_row`
4. **Dòng ngoài phạm vi**: LUÔN hiển thị (không bao giờ bị ẩn)

### Kiểm Tra Config

```python
from excel_automation import SizeFilterConfig

config = SizeFilterConfig()
is_valid, message = config.validate_config(max_row=100)

if is_valid:
    print("Config hợp lệ")
else:
    print(f"Lỗi: {message}")
```

## Lưu Ý Quan Trọng

### ⚠️ Giới Hạn Phạm Vi Ẩn Dòng

- Tính năng CHỈ ẩn/hiện dòng trong phạm vi `start_row` đến `end_row`
- Dòng ngoài phạm vi này KHÔNG BAO GIỜ bị ảnh hưởng
- Ví dụ: Nếu config là `19-59`, thì:
  - Dòng 1-18: LUÔN hiển thị
  - Dòng 19-59: Có thể ẩn/hiện
  - Dòng 60+: LUÔN hiển thị

### 📝 Mặc Định Unchecked

- Khi mở dialog lọc size, tất cả checkbox mặc định là **unchecked**
- Nghĩa là nếu không chọn gì, TẤT CẢ dòng sẽ bị ẩn
- Hãy chọn ít nhất 1 size trước khi áp dụng

### 💾 Lưu File

- Sau khi áp dụng filter, file Excel sẽ được lưu tự động
- Nên tải lại file trong Excel Viewer để xem kết quả
- Hoặc mở file bằng Excel để kiểm tra

## Troubleshooting

### Không tìm thấy size nào

**Nguyên nhân:**
- Cột không đúng
- Phạm vi dòng không đúng
- Dữ liệu không phải số 3 chữ số

**Giải pháp:**
- Kiểm tra lại cấu hình (Menu → Lọc Size → Cấu hình)
- Đảm bảo dữ liệu trong cột là số (044, 045...)

### Lỗi "vượt quá số dòng thực tế"

**Nguyên nhân:**
- `end_row` trong config lớn hơn số dòng thực tế trong sheet

**Giải pháp:**
- Mở dialog cấu hình
- Giảm `end_row` xuống phù hợp với số dòng thực tế

### Dòng ngoài phạm vi bị ẩn

**Không thể xảy ra:**
- Tính năng có validation chặt chẽ
- CHỈ ẩn dòng trong phạm vi `start_row` đến `end_row`
- Nếu gặp vấn đề này, vui lòng báo lỗi

## Demo Script

Chạy script demo để xem các ví dụ:

```bash
python scripts/size_filter_demo.py
```

## Unit Tests

Chạy tests để kiểm tra tính năng:

```bash
python tests/test_size_filter.py
```

Tất cả 11 tests phải pass.

