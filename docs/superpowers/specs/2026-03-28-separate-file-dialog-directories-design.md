# Separate File Dialog Directories

## Problem

Khi mở file Excel rồi mở file PDF (hoặc ngược lại), file dialog luôn quay về thư mục mặc định của hệ thống. Người dùng phải navigate lại thư mục mỗi lần chuyển loại file.

## Solution

Lưu thư mục cuối cùng cho 3 loại file riêng biệt vào `UIConfig`, sử dụng `initialdir` của `filedialog.askopenfilename`.

## Config

Thêm `last_directories` vào `UIConfig.DEFAULT_CONFIG`:

```python
"last_directories": {
    "excel_open": None,
    "pdf_import_po": None,
    "pdf_reader": None
}
```

Persist trong `config/ui_config.json`.

## UIConfig Methods

Thêm 2 method vào class `UIConfig` (`ui/ui_config.py`):

- `get_last_directory(key: str) -> Optional[str]`: Trả về đường dẫn thư mục từ `last_directories.<key>` nếu thư mục tồn tại trên disk. Trả `None` nếu chưa có hoặc thư mục đã bị xóa.
- `set_last_directory(key: str, file_path: str)`: Nhận đường dẫn file, extract thư mục cha (`Path(file_path).parent`), lưu vào `last_directories.<key>`.

## Files to Modify

### 1. `ui/ui_config.py`

- Thêm `"last_directories"` vào `DEFAULT_CONFIG`
- Thêm `get_last_directory()` và `set_last_directory()`

### 2. `ui/excel_realtime_controller.py` — `_open_file()`

- Trước dialog: `initialdir=self.config.get_last_directory("excel_open")`
- Sau chọn file: `self.config.set_last_directory("excel_open", file_path)`

### 3. `ui/excel_realtime_controller.py` — `_import_po_from_pdf()`

- Trước dialog: `initialdir=self.config.get_last_directory("pdf_import_po")`
- Sau chọn file: `self.config.set_last_directory("pdf_import_po", file_path)`

### 4. `ui/pdf_reader_dialog.py` — `_choose_and_read_pdf()`

- Thêm tham số `ui_config: Optional[UIConfig] = None` vào `PdfReaderDialog.__init__()`
- Trước dialog: `initialdir=self.ui_config.get_last_directory("pdf_reader")` nếu có config
- Sau chọn file: `self.ui_config.set_last_directory("pdf_reader", file_path)` nếu có config

### 5. `ui/excel_realtime_controller.py` — `_open_pdf_reader()`

- Truyền `self.config` vào `PdfReaderDialog(self.root, ui_config=self.config)`

## Behavior

- Lần đầu chạy: `initialdir=None` → tkinter mở thư mục mặc định (hành vi như cũ)
- Sau khi chọn file: thư mục cha được lưu vào config JSON
- Lần mở tiếp theo: dialog mở đúng thư mục đã lưu
- 3 loại file hoàn toàn độc lập, không ảnh hưởng lẫn nhau
- Nếu thư mục đã lưu bị xóa: fallback về `None` (hành vi mặc định)
