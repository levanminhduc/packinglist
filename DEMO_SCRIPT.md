# Demo Script - Excel Viewer với Validation

## Chuẩn Bị

```bash
# Đảm bảo có sample data
python scripts/create_sample_data.py

# Khởi động Excel Viewer
python excel_viewer.py
```

## Demo Flow

### Bước 1: Giới Thiệu Giao Diện

**Nói:**
"Đây là Excel Viewer - công cụ xem và validate file Excel. 
Giao diện gồm:
- Menu bar với File, Validation, Cài Đặt, Trợ Giúp
- Toolbar với các buttons thao tác nhanh
- Khu vực hiển thị sheet tabs
- Bảng dữ liệu chính
- Status bar ở dưới"

### Bước 2: Mở File Excel

**Thao tác:**
1. Click "📂 Mở File" hoặc nhấn Ctrl+O
2. Navigate đến `data/input/sample_orders.xlsx`
3. Click Open

**Nói:**
"Tôi sẽ mở file sample_orders.xlsx - file này chứa 25 dòng dữ liệu đơn hàng.
Trong đó có 15 dòng hợp lệ và 10 dòng có lỗi để demo validation."

**Kết quả:**
- File được load
- Hiển thị sheet "Orders"
- Status bar: "Đã tải: sample_orders.xlsx - 1 sheets"
- Row count: "25 dòng"

### Bước 3: Load Validation Rules

**Thao tác:**
1. Click "📋 Load Rules" hoặc nhấn Ctrl+L
2. Navigate đến `data/validation_rules/packing_list_rules.json`
3. Click Open

**Nói:**
"Bây giờ tôi sẽ load validation rules từ file JSON.
File này định nghĩa các quy tắc validation cho 8 cột:
- PO: Phải có format PO + 7 chữ số
- Style: Độ dài từ 3-20 ký tự
- Color: Bắt buộc
- Size: Phải là XS, S, M, L, XL, hoặc XXL
- Quantity: Số nguyên từ 1 đến 100,000
- ShipDate: Format YYYY-MM-DD
- Buyer: Bắt buộc
- Carton: Số nguyên từ 1 đến 10,000"

**Kết quả:**
- Validation label hiển thị: "📋 Rules: 8 cột, 16 rules" (màu xanh)
- Dialog thông báo: "Đã load 16 validation rules cho 8 cột"

### Bước 4: Validate Dữ Liệu

**Thao tác:**
1. Click "✓ Validate" hoặc nhấn Ctrl+V
2. Đợi validation hoàn thành

**Nói:**
"Giờ tôi sẽ validate dữ liệu. 
Validation engine sẽ kiểm tra từng dòng, từng cột theo các rules đã định nghĩa."

**Kết quả:**
- Status bar: "Đang validate dữ liệu..."
- Sau đó: "Validation hoàn thành: 11 lỗi"
- Validation label: "❌ Lỗi: 11/25" (màu đỏ)
- 10 dòng được highlight màu vàng trong bảng
- Dialog hỏi: "❌ Tìm thấy 11 lỗi trong 25 dòng. Bạn có muốn xem chi tiết không?"

### Bước 5: Xem Các Dòng Lỗi

**Thao tác:**
Scroll qua bảng, chỉ vào các dòng màu vàng

**Nói:**
"Các dòng có lỗi được highlight màu vàng với text màu đỏ.
Ví dụ:
- Dòng 12: PO = 'INVALID' - sai format
- Dòng 13: Style = 'AB' - quá ngắn
- Dòng 14: Color trống
- Dòng 15: Size = 'XXXL' - không hợp lệ
- Dòng 16: Quantity = -100 - số âm
- ..."

### Bước 6: Xem Chi Tiết Lỗi

**Thao tác:**
1. Click "Yes" trong dialog
2. Hoặc Menu → Validation → Xem Kết Quả Validation

**Nói:**
"Dialog kết quả validation hiển thị:

Phần Tổng Quan:
- Tổng số dòng: 25
- Dòng hợp lệ: 15
- Số lỗi: 11
- Trạng thái: ❌ FAIL

Phần Chi Tiết Lỗi:
Bảng với 5 cột cho mỗi lỗi:
- Dòng: Số dòng có lỗi
- Cột: Tên cột bị lỗi
- Giá Trị: Giá trị không hợp lệ
- Quy Tắc: Rule bị vi phạm
- Thông Báo Lỗi: Mô tả chi tiết"

**Thao tác:**
Scroll qua danh sách lỗi, chỉ vào một vài lỗi điển hình

### Bước 7: Export Báo Cáo Lỗi

**Thao tác:**
1. Click "Export Báo Cáo" trong dialog
2. Hoặc Menu → Validation → Export Báo Cáo Lỗi...
3. Chọn vị trí lưu: `data/output/error_report.xlsx`
4. Click Save

**Nói:**
"Tôi có thể export báo cáo lỗi ra file Excel.
File này sẽ có format đẹp với:
- Header màu đỏ
- Tất cả lỗi được liệt kê chi tiết
- Auto-adjust column width
- Borders cho tất cả cells

File này có thể gửi cho người nhập liệu để họ fix lỗi."

**Kết quả:**
- File Excel được tạo
- Dialog: "Đã export báo cáo lỗi tại: ..."

### Bước 8: Demo Keyboard Shortcuts

**Thao tác:**
1. Nhấn Ctrl+L → Load rules dialog mở
2. Cancel
3. Nhấn Ctrl+V → Validate ngay
4. Nhấn Ctrl+O → Open file dialog

**Nói:**
"Excel Viewer hỗ trợ keyboard shortcuts để thao tác nhanh:
- Ctrl+O: Mở file
- Ctrl+L: Load validation rules
- Ctrl+V: Validate dữ liệu
- Ctrl+Q: Thoát"

### Bước 9: Demo Clear Validation

**Thao tác:**
1. Menu → Validation → Xóa Validation
2. Click OK trong confirmation dialog

**Nói:**
"Nếu muốn validate lại với rules khác, 
tôi có thể xóa validation hiện tại.
Điều này sẽ:
- Xóa kết quả validation
- Xóa highlight màu vàng
- Reset validation label"

**Kết quả:**
- Highlight màu vàng biến mất
- Validation label trống
- Dialog: "Đã xóa validation"

### Bước 10: Demo Validate Sheet Khác

**Nói:**
"Nếu file có nhiều sheets, tôi có thể:
1. Load rules một lần
2. Click vào sheet tab khác
3. Validate sheet đó
4. Lặp lại cho các sheets khác"

**Thao tác:**
(Nếu có multi-sheet file, demo chuyển sheet và validate)

### Bước 11: Tổng Kết

**Nói:**
"Tóm lại, Excel Viewer với Validation giúp:

✅ Validate dữ liệu nhanh chóng
- Không cần viết code
- Chỉ cần load rules và click validate

✅ Visual feedback rõ ràng
- Highlight lỗi trực tiếp trong bảng
- Color coding dễ nhận biết

✅ Báo cáo chi tiết
- Dialog hiển thị đầy đủ thông tin
- Export Excel format chuyên nghiệp

✅ User-friendly
- Keyboard shortcuts tiện lợi
- Menu organization hợp lý
- Error handling tốt

Công cụ này rất hữu ích cho:
- QC dữ liệu trước khi import
- Review file từ đối tác
- Training về data quality
- Quick check file Excel"

## Q&A Scenarios

### Q1: "Tôi có thể tạo rules mới không?"

**A:** "Có, bạn tạo file JSON mới theo format:
```json
{
  "ColumnName": [
    {
      "type": "required",
      "error_message": "..."
    }
  ]
}
```
Sau đó load file đó vào Excel Viewer."

### Q2: "Validate có chậm không với file lớn?"

**A:** "Với file < 10,000 rows thì rất nhanh (< 1 giây).
File lớn hơn có thể mất vài giây.
Nếu cần validate file rất lớn, nên dùng script command-line."

### Q3: "Có thể validate nhiều files cùng lúc không?"

**A:** "Hiện tại validate từng file một.
Nhưng có thể:
1. Load rules một lần
2. Mở file 1 → Validate → Note kết quả
3. Mở file 2 → Validate → Note kết quả
4. ..."

### Q4: "Có thể fix lỗi trực tiếp trong Excel Viewer không?"

**A:** "Không, Excel Viewer chỉ để xem và validate.
Để fix lỗi:
1. Export báo cáo
2. Mở file gốc trong Excel
3. Fix theo báo cáo
4. Load lại file trong Excel Viewer
5. Validate lại"

### Q5: "Rules có thể validate cross-column không?"

**A:** "Có, sử dụng CustomRule với function.
Ví dụ: Kiểm tra Quantity phải nhỏ hơn Carton * 100."

## Demo Tips

1. **Chuẩn bị trước:**
   - Đảm bảo sample data có sẵn
   - Test run trước khi demo
   - Đóng các ứng dụng không cần thiết

2. **Trong khi demo:**
   - Nói chậm, rõ ràng
   - Chỉ vào các elements khi nói
   - Pause sau mỗi action để audience theo dõi
   - Highlight các features quan trọng

3. **Xử lý lỗi:**
   - Nếu có lỗi, giải thích calmly
   - Show error handling features
   - Restart nếu cần

4. **Kết thúc:**
   - Tóm tắt key points
   - Mở Q&A
   - Cung cấp tài liệu tham khảo

