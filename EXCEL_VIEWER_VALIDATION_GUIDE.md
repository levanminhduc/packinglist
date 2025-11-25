# Hướng Dẫn Sử Dụng Validation trong Excel Viewer

## Khởi động Excel Viewer

```bash
python excel_viewer.py
```

## Tính năng Validation mới

### 1. Load Validation Rules

**Cách 1: Sử dụng Menu**
- Menu → Validation → Load Rules từ JSON...
- Chọn file rules (mặc định: `data/validation_rules/packing_list_rules.json`)

**Cách 2: Sử dụng Toolbar**
- Click button "📋 Load Rules"

**Cách 3: Keyboard Shortcut**
- Nhấn `Ctrl+L`

**Kết quả:**
- Thanh toolbar hiển thị: "📋 Rules: X cột, Y rules" (màu xanh)
- Thông báo thành công

### 2. Validate Dữ Liệu

**Điều kiện:**
- Đã mở file Excel
- Đã load validation rules (nếu chưa, sẽ được hỏi)

**Cách 1: Sử dụng Menu**
- Menu → Validation → Validate Dữ Liệu

**Cách 2: Sử dụng Toolbar**
- Click button "✓ Validate"

**Cách 3: Keyboard Shortcut**
- Nhấn `Ctrl+V`

**Kết quả nếu PASS:**
- Thanh toolbar hiển thị: "✅ Valid: X dòng" (màu xanh)
- Thông báo: "✅ Tất cả X dòng dữ liệu đều hợp lệ!"

**Kết quả nếu FAIL:**
- Thanh toolbar hiển thị: "❌ Lỗi: X/Y" (màu đỏ)
- Các dòng có lỗi được highlight màu vàng trong bảng
- Dialog hỏi có muốn xem chi tiết không

### 3. Xem Kết Quả Validation

**Cách 1: Sau khi validate (nếu có lỗi)**
- Click "Yes" trong dialog

**Cách 2: Sử dụng Menu**
- Menu → Validation → Xem Kết Quả Validation

**Nội dung hiển thị:**

**Phần Tổng Quan:**
- Tổng số dòng
- Dòng hợp lệ
- Số lỗi
- Trạng thái (✅ PASS / ❌ FAIL)

**Phần Chi Tiết Lỗi (nếu có):**
Bảng với các cột:
- Dòng: Số dòng có lỗi
- Cột: Tên cột
- Giá Trị: Giá trị bị lỗi
- Quy Tắc: Rule bị vi phạm
- Thông Báo Lỗi: Mô tả chi tiết

### 4. Export Báo Cáo Lỗi

**Cách 1: Từ Dialog Kết Quả**
- Click button "Export Báo Cáo"

**Cách 2: Sử dụng Menu**
- Menu → Validation → Export Báo Cáo Lỗi...

**Kết quả:**
- File Excel được tạo với format đẹp
- Header màu đỏ
- Các cột: Dòng, Cột, Giá Trị, Quy Tắc Vi Phạm, Thông Báo Lỗi
- Auto-adjust column width
- Borders cho tất cả cells

### 5. Xóa Validation

**Sử dụng Menu:**
- Menu → Validation → Xóa Validation

**Kết quả:**
- Xóa validation result
- Xóa validator
- Xóa highlight màu vàng trong bảng
- Reset validation label

## Workflow Thực Tế

### Scenario 1: Validate file mới

```
1. Mở Excel Viewer
2. Mở file Excel (Ctrl+O)
3. Load validation rules (Ctrl+L)
   → Chọn: data/validation_rules/packing_list_rules.json
4. Validate dữ liệu (Ctrl+V)
5. Nếu có lỗi:
   - Xem các dòng highlight màu vàng
   - Xem chi tiết lỗi
   - Export báo cáo nếu cần
```

### Scenario 2: Validate nhiều sheets

```
1. Mở file Excel có nhiều sheets
2. Load validation rules (1 lần)
3. Click vào sheet tab để chuyển sheet
4. Validate sheet hiện tại (Ctrl+V)
5. Lặp lại bước 3-4 cho các sheets khác
```

### Scenario 3: Validate với rules khác nhau

```
1. Mở file Excel
2. Load rules set 1 (Ctrl+L)
3. Validate (Ctrl+V)
4. Xóa validation (Menu → Validation → Xóa Validation)
5. Load rules set 2 (Ctrl+L)
6. Validate lại (Ctrl+V)
```

## Visual Indicators

### Thanh Toolbar

**File Label:**
- "Chưa mở file nào" (gray) - Chưa mở file
- "📄 filename.xlsx" (black) - Đã mở file

**Validation Label:**
- "" (empty) - Chưa load rules
- "📋 Rules: X cột, Y rules" (blue) - Đã load rules
- "✅ Valid: X dòng" (green) - Validation pass
- "❌ Lỗi: X/Y" (red) - Validation fail

### Bảng Dữ Liệu

**Dòng bình thường:**
- Background: White
- Text: Black

**Dòng có lỗi:**
- Background: Yellow (#FFFF99)
- Text: Red (#CC0000)

### Status Bar

- "Sẵn sàng" - Idle
- "Đang đọc file..." - Loading
- "Đang validate dữ liệu..." - Validating
- "Validation hoàn thành: X lỗi" - Done

## Keyboard Shortcuts

| Shortcut | Chức năng |
|----------|-----------|
| Ctrl+O | Mở file Excel |
| Ctrl+L | Load validation rules |
| Ctrl+V | Validate dữ liệu |
| Ctrl+Q | Thoát |

## Tips & Tricks

### 1. Validate nhanh

Sau khi load rules lần đầu, chỉ cần:
- Mở file mới (Ctrl+O)
- Validate ngay (Ctrl+V)

### 2. So sánh trước/sau fix

1. Validate file gốc → Export báo cáo
2. Fix lỗi trong Excel
3. Tải lại file (🔄 button)
4. Validate lại
5. So sánh số lỗi

### 3. Batch validation

1. Load rules 1 lần
2. Mở file 1 → Validate → Ghi nhận kết quả
3. Mở file 2 → Validate → Ghi nhận kết quả
4. ...

### 4. Custom rules cho từng file

Tạo nhiều rules files:
- `packing_list_rules.json`
- `invoice_rules.json`
- `order_rules.json`

Load rules phù hợp với từng loại file

## Troubleshooting

### Lỗi: "Chưa mở file nào để validate"

**Nguyên nhân:** Chưa mở file Excel

**Giải pháp:** Mở file trước (Ctrl+O)

### Lỗi: "Chưa load validation rules"

**Nguyên nhân:** Chưa load rules file

**Giải pháp:** Load rules trước (Ctrl+L)

### Không thấy highlight màu vàng

**Nguyên nhân:** 
- Validation pass (không có lỗi)
- Đã xóa validation

**Giải pháp:** Validate lại (Ctrl+V)

### Export báo cáo bị lỗi

**Nguyên nhân:**
- Không có lỗi để export
- File đích đang mở

**Giải pháp:**
- Kiểm tra có lỗi không
- Đóng file Excel đích nếu đang mở

## File Paths Mặc Định

- **Validation Rules:** `data/validation_rules/`
- **Output Reports:** `data/output/`
- **Sample Data:** `data/input/sample_orders.xlsx`

## Ví Dụ Thực Tế

### Test với Sample Data

```bash
# 1. Chạy Excel Viewer
python excel_viewer.py

# 2. Trong Excel Viewer:
#    - Mở file: data/input/sample_orders.xlsx
#    - Load rules: data/validation_rules/packing_list_rules.json
#    - Click Validate
#    - Xem 10 dòng highlight màu vàng (có lỗi)
#    - Xem chi tiết 11 lỗi
#    - Export báo cáo
```

## Tích Hợp với Workflow

Excel Viewer với Validation có thể dùng để:

1. **QC dữ liệu trước khi import**
   - Validate file trước khi import vào database
   - Đảm bảo data quality

2. **Review dữ liệu từ partners**
   - Nhận file từ đối tác
   - Validate theo rules
   - Gửi lại báo cáo lỗi

3. **Training & Documentation**
   - Demo validation rules cho team
   - Giải thích các lỗi thường gặp

4. **Quick Check**
   - Kiểm tra nhanh file Excel
   - Không cần viết code

