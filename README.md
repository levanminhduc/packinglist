# Excel Automation - Dự Án Tự Động Hóa Excel

Dự án Python chuyên nghiệp để đọc, ghi và xử lý file Excel tự động trên máy local.

## 🎯 Tính Năng

### Excel Real-Time Controller (excel_realtime_controller.py)
Ứng dụng GUI điều khiển Excel real-time với các tính năng:

- 📂 **Mở File Excel**: Hỗ trợ .xlsx, .xls, .xlsm, .xlsb qua COM automation
- 📑 **Quản lý Sheets**: Chuyển đổi và reload sheets dễ dàng
- 🔍 **Quét Sizes**: Tự động quét và hiển thị danh sách sizes từ cột cấu hình
- 👁️ **Ẩn/Hiện Dòng**: Ẩn dòng real-time theo sizes đã chọn, hiện lại tất cả dòng
- 📝 **Update PO**: Cập nhật hàng loạt mã PO cho các dòng
- 🎨 **Update Color**: Cập nhật hàng loạt mã màu cho các dòng
- 📊 **Nhập Số Lượng**: Nhập số lượng cho từng size và ghi trực tiếp vào Excel
- ⚙️ **Cấu hình linh hoạt**: Tùy chỉnh cột, dòng bắt đầu/kết thúc để quét
- 💾 **Lưu vị trí cửa sổ**: Tự động nhớ vị trí và kích thước cửa sổ

### Excel Automation Core
- ✅ **Đọc Excel**: Hỗ trợ đọc file .xlsx, .xls, .xlsm, .xlsb
- ✅ **Ghi Excel**: Tạo và ghi file Excel với nhiều phương thức
- ✅ **Xử lý dữ liệu**: Làm sạch, lọc, tổng hợp, merge dữ liệu
- ✅ **Định dạng**: Tự động format header, borders, colors, freeze panes
- ✅ **Batch Processing**: Xử lý hàng loạt nhiều file
- ✅ **Backup tự động**: Tự động backup file trước khi xử lý
- ✅ **Logging**: Ghi log chi tiết mọi thao tác

## 📁 Cấu Trúc Dự Án

```
PythonExcel/
├── excel_automation/       # Package chính
│   ├── __init__.py
│   ├── reader.py          # Đọc Excel
│   ├── writer.py          # Ghi Excel
│   ├── processor.py       # Xử lý dữ liệu
│   ├── formatter.py       # Định dạng Excel
│   └── utils.py           # Tiện ích
├── config/                # Cấu hình
│   ├── __init__.py
│   └── settings.py
├── data/                  # Dữ liệu
│   ├── input/            # File đầu vào
│   ├── output/           # File đầu ra
│   ├── templates/        # Template Excel
│   └── backup/           # Backup files
├── scripts/              # Scripts automation
│   ├── daily_report.py   # Báo cáo hàng ngày
│   ├── data_import.py    # Import dữ liệu
│   └── batch_process.py  # Xử lý hàng loạt
├── tests/                # Unit tests
├── logs/                 # Log files
├── main.py              # Entry point
├── requirements.txt     # Dependencies
├── .env.example        # Environment template
└── README.md           # Tài liệu này
```

## 🚀 Cài Đặt

### 1. Clone hoặc tải dự án

```bash
cd PythonExcel
```

### 2. Tạo Virtual Environment (Khuyến nghị)

```bash
# Tạo virtual environment
python -m venv venv

# Kích hoạt (Windows)
venv\Scripts\activate

# Kích hoạt (Linux/macOS)
source venv/bin/activate
```

### 3. Cài đặt dependencies

```bash
pip install -r requirements.txt
```

### 4. Cấu hình môi trường

```bash
# Copy file .env.example thành .env
copy .env.example .env

# Chỉnh sửa .env theo nhu cầu (nếu cần)
```

## 📖 Hướng Dẫn Sử Dụng

### Chạy Excel Real-Time Controller

```bash
python excel_realtime_controller.py
```

**Quy trình sử dụng:**
1. Nhấn "📂 Chọn File Excel" để mở file
2. Chọn sheet cần làm việc từ dropdown
3. Nhấn "🔍 Quét Sizes" để tìm các sizes trong file
4. Chọn sizes cần hiển thị bằng checkbox
5. Sử dụng các chức năng:
   - **👁️ Ẩn dòng ngay**: Ẩn các dòng không thuộc sizes đã chọn
   - **👁️‍🗨️ Hiện tất cả**: Hiện lại tất cả dòng đã ẩn
   - **📝 Update PO**: Cập nhật mã PO hàng loạt
   - **🎨 Update Color**: Cập nhật mã màu hàng loạt
   - **📝 Nhập Số Lượng Size**: Nhập số lượng cho từng size
6. Nhấn "⚙️ Settings" để cấu hình cột và dòng quét

### Sử dụng qua Main.py

```bash
python main.py
```

### Chạy Scripts Riêng Lẻ

#### 1. Tạo báo cáo hàng ngày

```bash
python scripts/daily_report.py
```

#### 2. Import dữ liệu từ nhiều file

```bash
python scripts/data_import.py
```

#### 3. Xử lý hàng loạt

```bash
python scripts/batch_process.py
```

### Sử dụng trong Code Python

```python
from excel_automation import ExcelReader, ExcelWriter, ExcelProcessor, ExcelFormatter

# Đọc file Excel
reader = ExcelReader("data/input/myfile.xlsx")
df = reader.read_with_pandas()

# Xử lý dữ liệu
processor = ExcelProcessor()
df_clean = processor.clean_data(df, drop_duplicates=True)

# Ghi file Excel
writer = ExcelWriter("data/output/result.xlsx")
writer.write_dataframe(df_clean)

# Định dạng
formatter = ExcelFormatter("data/output/result.xlsx")
formatter.format_header()
formatter.auto_adjust_column_width()
```

## 🔧 Các Module Chính

### 1. ExcelReader - Đọc Excel

```python
reader = ExcelReader("file.xlsx")

# Đọc với pandas
df = reader.read_with_pandas(sheet_name="Sheet1")

# Đọc với openpyxl
ws = reader.read_with_openpyxl(sheet_name="Sheet1")

# Lấy danh sách sheets
sheets = reader.get_sheet_names()

# Đọc tất cả sheets
all_sheets = reader.read_all_sheets()
```

### 2. ExcelWriter - Ghi Excel

```python
writer = ExcelWriter("output.xlsx")

# Ghi DataFrame
writer.write_dataframe(df, sheet_name="Data")

# Ghi nhiều sheets
writer.write_multiple_sheets({
    'Sheet1': df1,
    'Sheet2': df2
})

# Append dữ liệu
writer.append_dataframe(df, sheet_name="Data")
```

### 3. ExcelProcessor - Xử lý dữ liệu

```python
processor = ExcelProcessor()

# Làm sạch dữ liệu
df_clean = processor.clean_data(df, drop_duplicates=True, fill_na=0)

# Lọc dữ liệu
df_filtered = processor.filter_data(df, {'Status': 'Active'})

# Tổng hợp
df_agg = processor.aggregate_data(df, group_by=['Category'], agg_dict={'Amount': 'sum'})

# Merge
df_merged = processor.merge_data(df1, df2, on='ID', how='inner')
```

### 4. ExcelFormatter - Định dạng

```python
formatter = ExcelFormatter("file.xlsx")

# Format header
formatter.format_header(bg_color="366092", font_color="FFFFFF")

# Tự động điều chỉnh độ rộng cột
formatter.auto_adjust_column_width()

# Thêm viền
formatter.add_borders()

# Freeze panes
formatter.freeze_panes(row=1)
```

## ⚙️ Cấu Hình

Chỉnh sửa file `.env` để thay đổi cấu hình:

```env
DEBUG=False
DATA_INPUT_DIR=data/input
DATA_OUTPUT_DIR=data/output
LOG_LEVEL=INFO
AUTO_BACKUP=True
BACKUP_KEEP_DAYS=30
```

## 📝 Logging

Tất cả hoạt động được ghi log vào `logs/app.log`:

```python
from excel_automation.utils import setup_logging
import logging

setup_logging("logs/app.log", logging.INFO)
logger = logging.getLogger(__name__)
logger.info("Thông báo của bạn")
```

## 🧪 Testing

```bash
# Chạy tests
pytest tests/

# Chạy với coverage
pytest tests/ --cov=excel_automation
```

## 🤝 Đóng Góp

Mọi đóng góp đều được hoan nghênh! Vui lòng:

1. Fork dự án
2. Tạo branch mới (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to branch (`git push origin feature/AmazingFeature`)
5. Tạo Pull Request

## 📄 License

Dự án này được phát hành dưới MIT License.

## 👤 Tác Giả

Your Name - your.email@example.com

## 🙏 Acknowledgments

- Pandas - Data manipulation
- OpenPyXL - Excel file handling
- XlsxWriter - Excel file creation

