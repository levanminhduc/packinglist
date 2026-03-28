# Separate File Dialog Directories — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Lưu thư mục cuối cùng riêng biệt cho 3 loại file dialog (Excel, PDF Import PO, PDF Reader) để người dùng không phải navigate lại mỗi lần chuyển loại file.

**Architecture:** Thêm `last_directories` vào `UIConfig.DEFAULT_CONFIG` với 3 key. Thêm 2 method tiện ích `get_last_directory()` / `set_last_directory()`. Sửa 3 chỗ gọi `filedialog.askopenfilename` để truyền `initialdir` và lưu thư mục sau khi chọn file.

**Tech Stack:** Python, tkinter filedialog, UIConfig JSON persistence

**Spec:** `docs/superpowers/specs/2026-03-28-separate-file-dialog-directories-design.md`

---

### Task 1: Thêm `last_directories` và 2 method vào UIConfig

**Files:**
- Modify: `ui/ui_config.py:12-36` (DEFAULT_CONFIG) và cuối class (thêm methods)

- [ ] **Step 1: Thêm `last_directories` vào DEFAULT_CONFIG**

Trong `ui/ui_config.py`, thêm key `"last_directories"` vào `DEFAULT_CONFIG` sau `"last_opened_file": None`:

```python
"last_directories": {
    "excel_open": None,
    "pdf_import_po": None,
    "pdf_reader": None
}
```

- [ ] **Step 2: Thêm method `get_last_directory`**

Thêm sau method `reset_to_defaults()` (line 135-138):

```python
def get_last_directory(self, key: str) -> Optional[str]:
    directory = self.get(f'last_directories.{key}')
    if directory and Path(directory).is_dir():
        return directory
    return None
```

Cần thêm `from pathlib import Path` vào đầu file (dòng 1-2).

- [ ] **Step 3: Thêm method `set_last_directory`**

Thêm ngay sau `get_last_directory`:

```python
def set_last_directory(self, key: str, file_path: str) -> None:
    directory = str(Path(file_path).parent)
    self.set(f'last_directories.{key}', directory)
```

- [ ] **Step 4: Commit**

```bash
git add ui/ui_config.py
git commit -m "feat: thêm last_directories vào UIConfig để nhớ thư mục file dialog"
```

---

### Task 2: Sửa `_open_file()` trong controller — nhớ thư mục Excel

**Files:**
- Modify: `ui/excel_realtime_controller.py:1-24` (imports) và `:32-36` (init) và `:320-327` (_open_file)

- [ ] **Step 1: Thêm import UIConfig**

Trong `ui/excel_realtime_controller.py`, thêm import sau dòng 23 (`from excel_automation.pdf_po_parser import PDFPOParser`):

```python
from ui.ui_config import UIConfig
```

- [ ] **Step 2: Tạo instance UIConfig trong `__init__`**

Thêm sau dòng `self.dialog_config = DialogConfigManager()` (line 35):

```python
self.ui_config = UIConfig()
```

- [ ] **Step 3: Sửa `_open_file()` — thêm `initialdir` và lưu thư mục**

Thay thế phần `filedialog.askopenfilename` trong `_open_file()` (line 320-331):

```python
def _open_file(self) -> None:
    file_path = filedialog.askopenfilename(
        title="Chọn File Excel",
        filetypes=[
            ("Excel Files", "*.xlsx *.xls *.xlsm *.xlsb"),
            ("All Files", "*.*")
        ],
        initialdir=self.ui_config.get_last_directory("excel_open")
    )

    if not file_path:
        return

    self.ui_config.set_last_directory("excel_open", file_path)

    try:
```

Phần `try:` trở đi giữ nguyên (line 332+).

- [ ] **Step 4: Commit**

```bash
git add ui/excel_realtime_controller.py
git commit -m "feat: nhớ thư mục cuối khi mở file Excel"
```

---

### Task 3: Sửa `_import_po_from_pdf()` — nhớ thư mục PDF Import PO

**Files:**
- Modify: `ui/excel_realtime_controller.py:1174-1184` (_import_po_from_pdf)

- [ ] **Step 1: Sửa `_import_po_from_pdf()` — thêm `initialdir` và lưu thư mục**

Thay thế phần `filedialog.askopenfilename` trong `_import_po_from_pdf()` (line 1179-1184):

```python
file_path = filedialog.askopenfilename(
    title="Chọn file PDF Purchase Order",
    filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    initialdir=self.ui_config.get_last_directory("pdf_import_po")
)
if not file_path:
    return

self.ui_config.set_last_directory("pdf_import_po", file_path)
```

Phần code từ `from ui.pdf_import_dialog import ...` (line 1186+) giữ nguyên.

- [ ] **Step 2: Commit**

```bash
git add ui/excel_realtime_controller.py
git commit -m "feat: nhớ thư mục cuối khi import PO từ PDF"
```

---

### Task 4: Sửa `PdfReaderDialog` — nhớ thư mục PDF Reader

**Files:**
- Modify: `ui/pdf_reader_dialog.py:21-28` (class init) và `:122-134` (_choose_and_read_pdf)
- Modify: `ui/excel_realtime_controller.py:1525-1529` (_open_pdf_reader)

- [ ] **Step 1: Thêm tham số `ui_config` vào `PdfReaderDialog.__init__`**

Sửa `__init__` (line 26-28):

```python
def __init__(self, parent: tk.Tk, ui_config=None):
    self.parent = parent
    self.ui_config = ui_config
    self.dialog_config = DialogConfigManager()
```

- [ ] **Step 2: Sửa `_choose_and_read_pdf()` — thêm `initialdir` và lưu thư mục**

Thay thế method (line 122-134):

```python
def _choose_and_read_pdf(self) -> None:
    initial_dir = self.ui_config.get_last_directory("pdf_reader") if self.ui_config else None
    file_path = filedialog.askopenfilename(
        title="Chọn file PDF",
        filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        parent=self.dialog,
        initialdir=initial_dir
    )

    if not file_path:
        return

    if self.ui_config:
        self.ui_config.set_last_directory("pdf_reader", file_path)

    self.file_label.config(text=file_path, foreground="black")
    self._start_extraction(file_path)
```

- [ ] **Step 3: Sửa `_open_pdf_reader()` trong controller — truyền `ui_config`**

Thay thế dòng `PdfReaderDialog(self.root)` (line 1529):

```python
PdfReaderDialog(self.root, ui_config=self.ui_config)
```

- [ ] **Step 4: Commit**

```bash
git add ui/pdf_reader_dialog.py ui/excel_realtime_controller.py
git commit -m "feat: nhớ thư mục cuối khi mở PDF Reader"
```

---

### Task 5: Test thủ công

- [ ] **Step 1: Chạy ứng dụng**

```bash
python excel_realtime_controller.py
```

- [ ] **Step 2: Test Excel file dialog**

Mở file Excel từ thư mục A → đóng dialog → mở lại → xác nhận dialog mở đúng thư mục A.

- [ ] **Step 3: Test PDF Import PO file dialog**

Mở PDF Import PO từ thư mục B → đóng dialog → mở lại → xác nhận dialog mở đúng thư mục B (khác thư mục A).

- [ ] **Step 4: Test PDF Reader file dialog**

Mở PDF Reader → chọn PDF từ thư mục C → đóng → mở lại PDF Reader → xác nhận dialog mở đúng thư mục C.

- [ ] **Step 5: Kiểm tra file config**

Mở `config/ui_config.json` → xác nhận có key `last_directories` với 3 đường dẫn riêng biệt.

- [ ] **Step 6: Test fallback**

Xóa key `last_directories` trong JSON → mở lại app → xác nhận dialog mở thư mục mặc định (không crash).
