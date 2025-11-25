# Quy trình nhập liệu tối ưu & vai trò của Python

## 1. Mục tiêu

- Giảm tối đa thao tác gõ tay, copy–paste lặp lại.
- Hạn chế lỗi nhập liệu (sai số, sai mã, trùng dữ liệu).
- Dễ kiểm soát, truy vết (log lại các lần nhập).
- Dễ mở rộng khi quy trình, biểu mẫu hoặc hệ thống thay đổi.

---

## 2. Quy trình làm việc tổng quát

### Bước 1: Nhận & tập hợp dữ liệu nguồn

Nguồn dữ liệu có thể là:
- File Excel/CSV từ khách hàng hoặc nội bộ.
- Biểu mẫu giấy được nhập lại vào Excel.
- Dữ liệu xuất từ hệ thống khác (ERP cũ, phần mềm kho, v.v.).

**Python có thể hỗ trợ:**
- Gộp nhiều file Excel/CSV thành 1 file tổng hợp.
- Đổi định dạng file (CSV ⇄ Excel, tách/shêp sheet, v.v.).
- Đọc tự động từ thư mục chỉ định (ví dụ: tất cả file trong `input/`).

---

### Bước 2: Làm sạch & chuẩn hóa dữ liệu

Các vấn đề thường gặp:
- Dữ liệu trùng (một PO nhập nhiều lần).
- Mã hàng/màu/size không thống nhất (chữ hoa/chữ thường, thừa khoảng trắng).
- Sai định dạng ngày tháng, số lượng không phải số.
- Thiếu thông tin bắt buộc (PO, Style, Color, Qty…).

**Python có thể hỗ trợ:**
- Dùng `pandas` để:
  - Xóa trùng (`drop_duplicates`).
  - Chuẩn hóa text (cắt khoảng trắng, viết hoa, chuẩn lại mã size).
  - Kiểm tra & loại các dòng thiếu dữ liệu bắt buộc.
  - Tính và đối chiếu tổng số lượng theo PO/Style/Color trước khi nhập hệ thống.
- Sinh file `clean_data.xlsx`/`error_data.xlsx` để người dùng kiểm tra.

---

### Bước 3: Tự động nhập liệu vào hệ thống

Các hệ thống đích có thể là:
- Website (web form) của khách hàng/đối tác.
- Phần mềm nội bộ chỉ có giao diện web.
- Cơ sở dữ liệu (MySQL, SQL Server, PostgreSQL...).
- API của ERP/WMS/MES.

**Python có thể hỗ trợ:**

1. **Nhập liệu qua trình duyệt (web form)**  
   - Dùng `selenium` hoặc `playwright`:
     - Mở trình duyệt (Chrome/Edge/Firefox).
     - Đăng nhập tự động (nếu được phép).
     - Lặp từng dòng dữ liệu:
       - Điền các ô input (PO, Style, Color, Qty…).
       - Chọn dropdown, tick checkbox, bấm nút `Save/Submit`.
   - Tự động dừng hoặc log lỗi nếu website trả về thông báo sai.

2. **Đẩy thẳng vào database hoặc API**
   - Dùng thư viện DB (`sqlalchemy`, `pyodbc`, `psycopg2`, v.v.) để:
     - Insert/update nhiều dòng một lúc.
     - Hạn chế phải nhập tay qua giao diện.
   - Dùng `requests` để gọi API:
     - Gửi dữ liệu dạng JSON theo format hệ thống yêu cầu.
     - Nhận response và ghi log kết quả.

---

### Bước 4: Kiểm tra, log & xử lý lỗi

Sau khi nhập liệu:

- Kiểm tra ngẫu nhiên một số record trên hệ thống so với file nguồn.
- Đối chiếu tổng số lượng (theo PO, Style, Color, Size) giữa:
  - File nguồn sạch (`clean_data`)  
  - Dữ liệu trên hệ thống (query từ DB hoặc báo cáo hệ thống).

**Python có thể hỗ trợ:**
- Ghi log từng dòng nhập (thời gian, PO, status OK/FAIL, thông báo lỗi).
- Tự sinh báo cáo “danh sách dòng lỗi” để người dùng xử lý.
- Tự chạy lại chỉ các dòng lỗi sau khi đã chỉnh sửa.

---

### Bước 5: Sinh báo cáo & tài liệu phục vụ sản xuất/vận chuyển

Sau khi dữ liệu đã vào hệ thống:

- Cần tạo các loại file:
  - Packing list, carton plan.
  - Phiếu giao hàng, invoice.
  - Báo cáo tổng hợp theo ngày/tuần/tháng, theo khách hàng hoặc PO.

**Python có thể hỗ trợ:**
- Dùng `pandas` để tổng hợp dữ liệu theo nhiều chiều.
- Dùng `openpyxl` hoặc `xlsxwriter` để xuất Excel với format sẵn (template).
- Tự động đặt tên file theo quy ước (ví dụ: `PACKINGLIST_PO9012689_2025-11-18.xlsx`).

---

## 3. Mô hình triển khai gợi ý

1. **Thư mục làm việc**
   - `input/` : chứa file nguồn gốc (Excel/CSV).
   - `output/` : chứa file sạch, file báo cáo, log.
   - `scripts/` : chứa các file `.py` xử lý từng bước.

2. **Các script Python chính**
   - `01_clean_data.py`  : đọc & chuẩn hóa dữ liệu.
   - `02_import_to_system.py` : tự động nhập vào web/DB/API.
   - `03_generate_reports.py` : sinh packing list, báo cáo tổng hợp.

3. **Cách sử dụng**
   - Người dùng chỉ cần:
     1. Copy file Excel nguồn vào `input/`.
     2. Chạy lần lượt:
        - `01_clean_data.py`
        - kiểm tra file sạch & file lỗi,
        - `02_import_to_system.py`,
        - `03_generate_reports.py` (nếu cần).
     3. Kiểm tra nhanh trên hệ thống & lưu trữ log.

---

## 4. Lợi ích khi áp dụng Python vào quy trình nhập liệu

- **Tiết kiệm thời gian:**  
  - Giảm đáng kể thao tác gõ tay lặp lại.
- **Giảm sai sót:**  
  - Logic kiểm tra nhất quán, không phụ thuộc tâm trạng/kinh nghiệm từng người.
- **Dễ mở rộng:**  
  - Khi mẫu Excel/web form thay đổi, chỉ cần sửa script thay vì đào tạo lại toàn bộ nhân viên.
- **Dễ audit & truy vết:**  
  - Log đầy đủ, dễ xem lại “ai nhập, nhập lúc nào, dòng nào lỗi”.

---

## 5. Hướng phát triển tiếp theo

- Viết thêm giao diện đơn giản (web app bằng Flask/Streamlit) để người không biết Python vẫn bấm nút chạy được.
- Kết hợp Task Scheduler (Windows) hoặc cron (Linux) để:
  - Tự động chạy script vào khung giờ cố định mỗi ngày.
- Tích hợp với hệ thống quyền & phân quyền nếu nhiều phòng ban cùng dùng.