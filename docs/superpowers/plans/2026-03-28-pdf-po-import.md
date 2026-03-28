# Import PO từ PDF — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Thêm tính năng Import PO từ PDF — tự động trích xuất PO, Color, Sizes, Quantities từ file PDF Purchase Order và ghi vào Excel đang mở, với preview dialog cho user review/sửa trước khi ghi.

**Architecture:** Tạo module `PDFPOParser` (business logic parse PDF) + `PDFImportDialog` (preview + progress UI). Tích hợp vào main UI bằng nút mới đứng đầu tiên với style xanh lá. Tái sử dụng `POUpdateManager` và `ColorCodeUpdateManager` để ghi Excel.

**Tech Stack:** Python 3, pdfplumber, tkinter, win32com (COM), pytest

**Spec:** `docs/superpowers/specs/2026-03-28-pdf-po-import-design.md`

---

## File Structure

| Action | File | Responsibility |
|---|---|---|
| Create | `excel_automation/pdf_po_parser.py` | Parse PDF → trích xuất PO, Color, Sizes, Qty |
| Create | `ui/pdf_import_dialog.py` | Preview dialog + Progress dialog |
| Create | `tests/test_pdf_po_parser.py` | Unit tests cho parser |
| Modify | `ui/excel_realtime_controller.py` | Thêm nút Import, ẩn Quét Sizes, thêm method `_import_po_from_pdf` |
| Modify | `data/template_configs/dialog_config.json` | Thêm kích thước dialog mới |

---

### Task 1: PDFPOData dataclass

**Files:**
- Create: `excel_automation/pdf_po_parser.py`
- Create: `tests/test_pdf_po_parser.py`

- [ ] **Step 1: Tạo file parser với dataclass**

```python
# excel_automation/pdf_po_parser.py
from dataclasses import dataclass, field
from typing import Dict
import logging

logger = logging.getLogger(__name__)


@dataclass
class PDFPOData:
    raw_po: str
    po_number: str
    color_code: str
    size_quantities: Dict[str, int] = field(default_factory=dict)
    total_quantity: int = 0
    source_file: str = ""
```

- [ ] **Step 2: Viết test cho dataclass**

```python
# tests/test_pdf_po_parser.py
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pytest
from excel_automation.pdf_po_parser import PDFPOData


class TestPDFPOData:

    def test_create_pdf_po_data(self):
        data = PDFPOData(
            raw_po="0009013330-1",
            po_number="9013330",
            color_code="3104",
            size_quantities={"046": 60, "048": 140},
            total_quantity=200,
            source_file="Test.pdf"
        )
        assert data.po_number == "9013330"
        assert data.color_code == "3104"
        assert data.size_quantities["046"] == 60
        assert data.total_quantity == 200

    def test_default_values(self):
        data = PDFPOData(raw_po="", po_number="", color_code="")
        assert data.size_quantities == {}
        assert data.total_quantity == 0
        assert data.source_file == ""
```

- [ ] **Step 3: Chạy test xác nhận pass**

Run: `pytest tests/test_pdf_po_parser.py -v`
Expected: 2 tests PASS

- [ ] **Step 4: Commit**

```bash
git add excel_automation/pdf_po_parser.py tests/test_pdf_po_parser.py
git commit -m "feat: add PDFPOData dataclass for PDF PO import"
```

---

### Task 2: PO number extraction

**Files:**
- Modify: `excel_automation/pdf_po_parser.py`
- Modify: `tests/test_pdf_po_parser.py`

- [ ] **Step 1: Viết failing tests cho PO extraction**

Thêm vào `tests/test_pdf_po_parser.py`:

```python
from excel_automation.pdf_po_parser import PDFPOData, PDFPOParser


class TestPOExtraction:

    def test_extract_po_standard(self):
        text = "P. O. No. Your reference\n0009013330-1 Marina Scholander"
        result = PDFPOParser._extract_po_number(text)
        assert result == ("0009013330-1", "9013330")

    def test_extract_po_strip_leading_zeros(self):
        text = "P. O. No. Your reference\n0009013330-1 Someone"
        raw, cleaned = PDFPOParser._extract_po_number(text)
        assert raw == "0009013330-1"
        assert cleaned == "9013330"

    def test_extract_po_no_leading_zeros(self):
        text = "P. O. No. Your reference\n9013330-2 Someone"
        raw, cleaned = PDFPOParser._extract_po_number(text)
        assert raw == "9013330-2"
        assert cleaned == "9013330"

    def test_extract_po_not_found(self):
        text = "This text has no PO number"
        with pytest.raises(RuntimeError, match="Không tìm thấy PO Number"):
            PDFPOParser._extract_po_number(text)
```

- [ ] **Step 2: Chạy test xác nhận fail**

Run: `pytest tests/test_pdf_po_parser.py::TestPOExtraction -v`
Expected: FAIL — `PDFPOParser` chưa tồn tại

- [ ] **Step 3: Implement PO extraction**

Thêm vào `excel_automation/pdf_po_parser.py`:

```python
import re


class PDFPOParser:

    @staticmethod
    def _extract_po_number(full_text: str) -> tuple:
        pattern = r'P\.?\s*O\.?\s*No\.?\s*[^\n]*\n\s*(\S+)'
        match = re.search(pattern, full_text)
        if not match:
            raise RuntimeError("Không tìm thấy PO Number trong file PDF")

        raw_po = match.group(1).strip()
        po_part = raw_po.split('-')[0]
        cleaned = po_part.lstrip('0') or '0'
        return (raw_po, cleaned)
```

- [ ] **Step 4: Chạy test xác nhận pass**

Run: `pytest tests/test_pdf_po_parser.py::TestPOExtraction -v`
Expected: 4 tests PASS

- [ ] **Step 5: Commit**

```bash
git add excel_automation/pdf_po_parser.py tests/test_pdf_po_parser.py
git commit -m "feat: add PO number extraction from PDF text"
```

---

### Task 3: Color code extraction

**Files:**
- Modify: `excel_automation/pdf_po_parser.py`
- Modify: `tests/test_pdf_po_parser.py`

- [ ] **Step 1: Viết failing tests cho Color extraction**

Thêm vào `tests/test_pdf_po_parser.py`:

```python
class TestColorExtraction:

    def test_extract_color_from_article_no(self):
        text = "000010 62183104046 AW Stretch Trousers 60 20.290 1217.40 USD"
        result = PDFPOParser._extract_color_code(text)
        assert result == "3104"

    def test_extract_color_different_article(self):
        text = "000010 62189999046 Some Product 60 20.290 1217.40 USD"
        result = PDFPOParser._extract_color_code(text)
        assert result == "9999"

    def test_extract_color_not_found(self):
        text = "No article numbers here"
        with pytest.raises(RuntimeError, match="Không tìm thấy Article Number"):
            PDFPOParser._extract_color_code(text)
```

- [ ] **Step 2: Chạy test xác nhận fail**

Run: `pytest tests/test_pdf_po_parser.py::TestColorExtraction -v`
Expected: FAIL — `_extract_color_code` chưa tồn tại

- [ ] **Step 3: Implement Color extraction**

Thêm vào class `PDFPOParser` trong `excel_automation/pdf_po_parser.py`:

```python
    @staticmethod
    def _extract_color_code(full_text: str) -> str:
        pattern = r'(?:^|\s)(\d{11,})\s'
        match = re.search(pattern, full_text)
        if not match:
            raise RuntimeError("Không tìm thấy Article Number trong file PDF")

        article_no = match.group(1)
        color_code = article_no[4:8]
        return color_code
```

- [ ] **Step 4: Chạy test xác nhận pass**

Run: `pytest tests/test_pdf_po_parser.py::TestColorExtraction -v`
Expected: 3 tests PASS

- [ ] **Step 5: Commit**

```bash
git add excel_automation/pdf_po_parser.py tests/test_pdf_po_parser.py
git commit -m "feat: add color code extraction from Article Number"
```

---

### Task 4: Size + Quantity extraction with normalization

**Files:**
- Modify: `excel_automation/pdf_po_parser.py`
- Modify: `tests/test_pdf_po_parser.py`

- [ ] **Step 1: Viết failing tests cho Size/Qty extraction**

Thêm vào `tests/test_pdf_po_parser.py`:

```python
class TestSizeQuantityExtraction:

    def test_extract_single_size_qty(self):
        text = "000010 62183104046 AW Stretch Trousers 60 20.290 1217.40 USD\nSize:46"
        result = PDFPOParser._extract_size_quantities(text)
        assert result == {"046": 60}

    def test_extract_multiple_sizes(self):
        text = (
            "000010 62183104046 AW Stretch Trousers 60 20.290 1217.40 USD\nSize:46\n"
            "000020 62183104048 AW Stretch Trousers 140 20.290 2840.60 USD\nSize:48\n"
            "000030 62183104050 AW Stretch Trousers 200 20.290 4058.00 USD\nSize:50"
        )
        result = PDFPOParser._extract_size_quantities(text)
        assert result == {"046": 60, "048": 140, "050": 200}

    def test_normalize_size_below_100(self):
        text = "000010 62183104096 AW Stretch Trousers 20 20.290 405.80 USD\nSize:96"
        result = PDFPOParser._extract_size_quantities(text)
        assert result == {"096": 20}

    def test_normalize_size_100_and_above(self):
        text = "000010 62183104100 AW Stretch Trousers 20 20.290 405.80 USD\nSize:100"
        result = PDFPOParser._extract_size_quantities(text)
        assert result == {"100": 20}

    def test_normalize_size_large(self):
        text = "000010 62183104148 AW Stretch Trousers 20 20.290 405.80 USD\nSize:148"
        result = PDFPOParser._extract_size_quantities(text)
        assert result == {"148": 20}

    def test_no_sizes_found(self):
        text = "No sizes here"
        with pytest.raises(RuntimeError, match="Không tìm thấy dữ liệu Size"):
            PDFPOParser._extract_size_quantities(text)
```

- [ ] **Step 2: Chạy test xác nhận fail**

Run: `pytest tests/test_pdf_po_parser.py::TestSizeQuantityExtraction -v`
Expected: FAIL — `_extract_size_quantities` chưa tồn tại

- [ ] **Step 3: Implement Size/Qty extraction**

Thêm vào class `PDFPOParser` trong `excel_automation/pdf_po_parser.py`:

```python
    @staticmethod
    def _normalize_size(size_str: str) -> str:
        size_str = size_str.strip()
        try:
            size_num = int(size_str)
            return str(size_num).zfill(3)
        except ValueError:
            return size_str

    @staticmethod
    def _extract_size_quantities(full_text: str) -> Dict[str, int]:
        lines = full_text.split('\n')
        size_quantities: Dict[str, int] = {}

        current_qty = None
        for line in lines:
            line = line.strip()

            qty_match = re.match(
                r'^0{2,}\d+\s+\d{11,}\s+.*?\s+(\d+)\s+[\d.]+\s+[\d,.]+\s+\w{3}$',
                line
            )
            if qty_match:
                current_qty = int(qty_match.group(1))
                continue

            size_match = re.match(r'^Size:\s*(.+)$', line)
            if size_match and current_qty is not None:
                raw_size = size_match.group(1).strip()
                normalized = PDFPOParser._normalize_size(raw_size)
                size_quantities[normalized] = current_qty
                current_qty = None

        if not size_quantities:
            raise RuntimeError("Không tìm thấy dữ liệu Size/Quantity trong file PDF")

        return size_quantities
```

- [ ] **Step 4: Chạy test xác nhận pass**

Run: `pytest tests/test_pdf_po_parser.py::TestSizeQuantityExtraction -v`
Expected: 6 tests PASS

- [ ] **Step 5: Commit**

```bash
git add excel_automation/pdf_po_parser.py tests/test_pdf_po_parser.py
git commit -m "feat: add size/quantity extraction with 3-digit normalization"
```

---

### Task 5: Full parse method + integration test với Test.pdf

**Files:**
- Modify: `excel_automation/pdf_po_parser.py`
- Modify: `tests/test_pdf_po_parser.py`

- [ ] **Step 1: Viết failing test cho full parse**

Thêm vào `tests/test_pdf_po_parser.py`:

```python
from pathlib import Path as TestPath


class TestFullParse:

    def test_parse_test_pdf(self):
        pdf_path = TestPath(__file__).parent.parent / "Test.pdf"
        if not pdf_path.exists():
            pytest.skip("Test.pdf không tồn tại")

        result = PDFPOParser.parse(str(pdf_path))

        assert isinstance(result, PDFPOData)
        assert result.raw_po == "0009013330-1"
        assert result.po_number == "9013330"
        assert result.color_code == "3104"
        assert result.total_quantity == 1030
        assert result.size_quantities["046"] == 60
        assert result.size_quantities["048"] == 140
        assert result.size_quantities["050"] == 200
        assert result.size_quantities["052"] == 200
        assert result.size_quantities["054"] == 160
        assert result.size_quantities["056"] == 100
        assert result.size_quantities["058"] == 20
        assert result.size_quantities["096"] == 20
        assert result.size_quantities["100"] == 20
        assert result.size_quantities["104"] == 20
        assert result.size_quantities["108"] == 20
        assert result.size_quantities["120"] == 10
        assert result.size_quantities["148"] == 20
        assert result.size_quantities["150"] == 20
        assert result.size_quantities["152"] == 20
        assert len(result.size_quantities) == 15
        assert result.source_file == str(pdf_path)

    def test_parse_nonexistent_file(self):
        with pytest.raises(RuntimeError, match="Không thể đọc file PDF"):
            PDFPOParser.parse("nonexistent.pdf")
```

- [ ] **Step 2: Chạy test xác nhận fail**

Run: `pytest tests/test_pdf_po_parser.py::TestFullParse -v`
Expected: FAIL — `PDFPOParser.parse` chưa tồn tại

- [ ] **Step 3: Implement full parse method**

Thêm vào class `PDFPOParser` trong `excel_automation/pdf_po_parser.py`:

```python
import pdfplumber
from pathlib import Path

    @staticmethod
    def parse(file_path: str) -> PDFPOData:
        path = Path(file_path)
        if not path.exists():
            raise RuntimeError(f"Không thể đọc file PDF: File không tồn tại: {file_path}")

        try:
            full_text = ""
            with pdfplumber.open(str(path)) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        full_text += page_text + "\n"
        except Exception as e:
            logger.error(f"Lỗi khi đọc file PDF: {e}")
            raise RuntimeError(f"Không thể đọc file PDF: {str(e)}")

        if not full_text.strip():
            raise RuntimeError("File PDF không chứa nội dung text")

        raw_po, po_number = PDFPOParser._extract_po_number(full_text)
        color_code = PDFPOParser._extract_color_code(full_text)
        size_quantities = PDFPOParser._extract_size_quantities(full_text)
        total_quantity = sum(size_quantities.values())

        logger.info(
            f"Parse PDF thành công: PO={po_number}, Color={color_code}, "
            f"{len(size_quantities)} sizes, total={total_quantity}"
        )

        return PDFPOData(
            raw_po=raw_po,
            po_number=po_number,
            color_code=color_code,
            size_quantities=size_quantities,
            total_quantity=total_quantity,
            source_file=str(path)
        )
```

- [ ] **Step 4: Chạy test xác nhận pass**

Run: `pytest tests/test_pdf_po_parser.py::TestFullParse -v`
Expected: 2 tests PASS (hoặc 1 skip nếu Test.pdf không có)

- [ ] **Step 5: Chạy toàn bộ test file**

Run: `pytest tests/test_pdf_po_parser.py -v`
Expected: Tất cả PASS

- [ ] **Step 6: Commit**

```bash
git add excel_automation/pdf_po_parser.py tests/test_pdf_po_parser.py
git commit -m "feat: add full PDF parse method with pdfplumber integration"
```

---

### Task 6: PDFImportDialog — Preview dialog

**Files:**
- Create: `ui/pdf_import_dialog.py`

- [ ] **Step 1: Tạo file với class PDFImportDialog**

```python
# ui/pdf_import_dialog.py
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, Dict, List, Optional
import logging

from excel_automation.pdf_po_parser import PDFPOData
from excel_automation.dialog_config_manager import DialogConfigManager

logger = logging.getLogger(__name__)


class PDFImportDialog:

    def __init__(
        self,
        parent: tk.Tk,
        pdf_data: PDFPOData,
        available_sizes: List[str],
        on_confirm_callback: Callable[[str, str, Dict[str, int]], None]
    ):
        self.parent = parent
        self.pdf_data = pdf_data
        self.available_sizes = available_sizes
        self.on_confirm_callback = on_confirm_callback
        self.dialog_config = DialogConfigManager()

        self.size_checkboxes: Dict[str, tk.BooleanVar] = {}
        self.size_entries: Dict[str, ttk.Entry] = {}
        self.status_labels: Dict[str, ttk.Label] = {}

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("📄 Import PO từ PDF")

        width, height = self.dialog_config.get_dialog_size('pdf_import')
        self.dialog.geometry(f"{width}x{height}")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_closing)

        self._create_widgets()
        self._center_window()

    def _center_window(self) -> None:
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f"+{x}+{y}")

    def _create_widgets(self) -> None:
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_label = ttk.Label(
            main_frame,
            text=f"File: {self.pdf_data.source_file}",
            foreground="gray"
        )
        file_label.pack(anchor=tk.W, pady=(0, 10))

        info_frame = ttk.LabelFrame(main_frame, text="Thông tin PO", padding=8)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        po_row = ttk.Frame(info_frame)
        po_row.pack(fill=tk.X, pady=2)
        ttk.Label(po_row, text="PO Number:", width=12).pack(side=tk.LEFT)
        self.po_var = tk.StringVar(value=self.pdf_data.po_number)
        ttk.Entry(po_row, textvariable=self.po_var, width=20, font=("Consolas", 11, "bold")).pack(side=tk.LEFT, padx=(5, 0))

        color_row = ttk.Frame(info_frame)
        color_row.pack(fill=tk.X, pady=2)
        ttk.Label(color_row, text="Color Code:", width=12).pack(side=tk.LEFT)
        self.color_var = tk.StringVar(value=self.pdf_data.color_code)
        ttk.Entry(color_row, textvariable=self.color_var, width=20, font=("Consolas", 11, "bold")).pack(side=tk.LEFT, padx=(5, 0))

        total_row = ttk.Frame(info_frame)
        total_row.pack(fill=tk.X, pady=2)
        ttk.Label(total_row, text="Total Qty:", width=12).pack(side=tk.LEFT)
        ttk.Label(total_row, text=f"{self.pdf_data.total_quantity:,}", font=("Consolas", 11, "bold"), foreground="#e65100").pack(side=tk.LEFT, padx=(5, 0))

        self._create_size_table(main_frame)
        self._create_warning_section(main_frame)
        self._create_buttons(main_frame)

    def _create_size_table(self, parent: ttk.Frame) -> None:
        size_frame = ttk.LabelFrame(parent, text="📋 Chi tiết Size — Quantity", padding=8)
        size_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        canvas = tk.Canvas(size_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(size_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable = ttk.Frame(canvas)

        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        header = ttk.Frame(scrollable)
        header.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header, text="☑", width=3, font=("", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="Size", width=10, font=("", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="Qty", width=8, font=("", 9, "bold"), anchor=tk.E).pack(side=tk.LEFT)
        ttk.Label(header, text="Trạng thái", width=16, font=("", 9, "bold")).pack(side=tk.LEFT, padx=(10, 0))

        for size, qty in self.pdf_data.size_quantities.items():
            is_match = size in self.available_sizes
            row = ttk.Frame(scrollable)
            row.pack(fill=tk.X, pady=1)

            var = tk.BooleanVar(value=is_match)
            self.size_checkboxes[size] = var
            ttk.Checkbutton(row, variable=var, command=lambda s=size: self._on_check_changed(s)).pack(side=tk.LEFT)

            size_entry = ttk.Entry(row, width=10, font=("Consolas", 10))
            size_entry.insert(0, size)
            if is_match:
                size_entry.configure(state="readonly")
            else:
                size_entry.bind('<KeyRelease>', lambda e, s=size: self._on_size_edited(s))
            size_entry.pack(side=tk.LEFT, padx=(2, 0))
            self.size_entries[size] = size_entry

            ttk.Label(row, text=str(qty), width=8, anchor=tk.E, font=("Consolas", 10)).pack(side=tk.LEFT)

            status_text = "✅ Khớp" if is_match else "⚠️ Chỉ có trong PDF"
            status_fg = "#2e7d32" if is_match else "#c62828"
            status_label = ttk.Label(row, text=status_text, foreground=status_fg, width=18)
            status_label.pack(side=tk.LEFT, padx=(10, 0))
            self.status_labels[size] = status_label

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_warning_section(self, parent: ttk.Frame) -> None:
        pdf_only = [s for s in self.pdf_data.size_quantities if s not in self.available_sizes]
        excel_only = [s for s in self.available_sizes if s not in self.pdf_data.size_quantities]

        if not pdf_only and not excel_only:
            return

        warn_frame = ttk.Frame(parent)
        warn_frame.pack(fill=tk.X, pady=(0, 10))

        if pdf_only:
            ttk.Label(
                warn_frame,
                text=f"⚠️ {len(pdf_only)} size chỉ có trong PDF: {', '.join(pdf_only)}",
                foreground="#c62828"
            ).pack(anchor=tk.W)

        if excel_only:
            ttk.Label(
                warn_frame,
                text=f"ℹ️ {len(excel_only)} size chỉ có trong Excel: {', '.join(excel_only)}",
                foreground="#1565c0"
            ).pack(anchor=tk.W)

    def _create_buttons(self, parent: ttk.Frame) -> None:
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Hủy", command=self._on_closing, width=15).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="✅ Xác nhận & Ghi", command=self._on_confirm, width=20).pack(side=tk.RIGHT, padx=(0, 5))

    def _on_size_edited(self, original_size: str) -> None:
        entry = self.size_entries[original_size]
        new_value = entry.get().strip()
        label = self.status_labels[original_size]

        if new_value in self.available_sizes:
            label.configure(text="✅ Khớp", foreground="#2e7d32")
            self.size_checkboxes[original_size].set(True)
        else:
            label.configure(text="⚠️ Chỉ có trong PDF", foreground="#c62828")

    def _on_check_changed(self, size: str) -> None:
        pass

    def _on_confirm(self) -> None:
        po = self.po_var.get().strip()
        color = self.color_var.get().strip()

        if not po:
            messagebox.showerror("Lỗi", "PO Number không được để trống", parent=self.dialog)
            return
        if not color:
            messagebox.showerror("Lỗi", "Color Code không được để trống", parent=self.dialog)
            return

        confirmed_sizes: Dict[str, int] = {}
        for original_size, var in self.size_checkboxes.items():
            if var.get():
                entry = self.size_entries[original_size]
                actual_size = entry.get().strip()
                qty = self.pdf_data.size_quantities[original_size]
                confirmed_sizes[actual_size] = qty

        if not confirmed_sizes:
            messagebox.showwarning("Cảnh báo", "Chưa chọn size nào để import", parent=self.dialog)
            return

        total = sum(confirmed_sizes.values())
        confirm_msg = (
            f"Xác nhận import vào Excel:\n\n"
            f"PO: {po}\n"
            f"Color: {color}\n"
            f"Sizes: {len(confirmed_sizes)} size\n"
            f"Total Qty: {total:,}\n\n"
            f"Tiếp tục?"
        )
        if messagebox.askyesno("Xác nhận Import", confirm_msg, parent=self.dialog):
            self._save_size_and_close()
            self.on_confirm_callback(po, color, confirmed_sizes)

    def _on_closing(self) -> None:
        self._save_size_and_close()

    def _save_size_and_close(self) -> None:
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            self.dialog_config.save_dialog_size('pdf_import', width, height)
        except Exception as e:
            logger.error(f"Lỗi khi lưu kích thước dialog: {e}")
        self.dialog.destroy()
```

- [ ] **Step 2: Commit**

```bash
git add ui/pdf_import_dialog.py
git commit -m "feat: add PDF import preview dialog with editable sizes"
```

---

### Task 7: ImportProgressDialog — Progress UI

**Files:**
- Modify: `ui/pdf_import_dialog.py`

- [ ] **Step 1: Thêm class ImportProgressDialog vào `ui/pdf_import_dialog.py`**

```python
class ImportProgressDialog:

    STEPS = [
        "Đọc file PDF",
        "Trích xuất dữ liệu PO, Color, Sizes",
        "Scan sizes từ Excel",
        "Ghi PO vào Excel",
        "Ghi Color Code vào Excel",
        "Cập nhật Sizes & Quantities",
        "Hoàn tất",
    ]

    STEP_WEIGHTS = [20, 15, 15, 15, 15, 15, 5]

    def __init__(self, parent: tk.Tk):
        self.parent = parent
        self.current_step = 0
        self.step_labels: List[ttk.Label] = []
        self.retry_callback: Optional[Callable] = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("📄 Đang Import PO từ PDF")
        self.dialog.geometry("420x380")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)

        self._create_widgets()
        self._center_window()

    def _center_window(self) -> None:
        self.dialog.update_idletasks()
        w = self.dialog.winfo_width()
        h = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (w // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (h // 2)
        self.dialog.geometry(f"+{x}+{y}")

    def _create_widgets(self) -> None:
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(
            main_frame, variable=self.progress_var,
            maximum=100, length=350, mode='determinate'
        )
        self.progress_bar.pack(pady=(0, 5))

        self.percent_label = ttk.Label(main_frame, text="0%", font=("", 11, "bold"))
        self.percent_label.pack(pady=(0, 15))

        steps_frame = ttk.Frame(main_frame)
        steps_frame.pack(fill=tk.BOTH, expand=True)

        for step_text in self.STEPS:
            label = ttk.Label(steps_frame, text=f"  ⬚  {step_text}", foreground="gray")
            label.pack(anchor=tk.W, pady=2)
            self.step_labels.append(label)

        self.btn_frame = ttk.Frame(main_frame)
        self.btn_frame.pack(fill=tk.X, pady=(15, 0))

        self.error_label = ttk.Label(main_frame, text="", foreground="#c62828", wraplength=360)

    def start_step(self, step_index: int) -> None:
        self.current_step = step_index
        percent = sum(self.STEP_WEIGHTS[:step_index])
        self.progress_var.set(percent)
        self.percent_label.configure(text=f"{percent}%")

        for i, label in enumerate(self.step_labels):
            if i < step_index:
                label.configure(text=f"  ✅  {self.STEPS[i]}", foreground="#2e7d32")
            elif i == step_index:
                label.configure(text=f"  🔄  Đang {self.STEPS[i].lower()}...", foreground="#1565c0")
            else:
                label.configure(text=f"  ⬚  {self.STEPS[i]}", foreground="gray")

        self.dialog.update()

    def complete_step(self, step_index: int) -> None:
        self.step_labels[step_index].configure(
            text=f"  ✅  {self.STEPS[step_index]}", foreground="#2e7d32"
        )
        percent = sum(self.STEP_WEIGHTS[:step_index + 1])
        self.progress_var.set(percent)
        self.percent_label.configure(text=f"{percent}%")
        self.dialog.update()

    def finish(self) -> None:
        self.progress_var.set(100)
        self.percent_label.configure(text="100%")
        for i, label in enumerate(self.step_labels):
            label.configure(text=f"  ✅  {self.STEPS[i]}", foreground="#2e7d32")
        self.dialog.update()
        self.parent.after(1000, self.dialog.destroy)

    def show_error(self, step_index: int, error_msg: str, retry_callback: Callable) -> None:
        self.step_labels[step_index].configure(
            text=f"  ❌  {self.STEPS[step_index]}", foreground="#c62828"
        )
        self.error_label.configure(text=f"Lỗi: {error_msg}")
        self.error_label.pack(pady=(10, 0))

        self.retry_callback = retry_callback

        for widget in self.btn_frame.winfo_children():
            widget.destroy()

        ttk.Button(
            self.btn_frame, text="🔄 Thử lại",
            command=self._retry, width=15
        ).pack(side=tk.LEFT)
        ttk.Button(
            self.btn_frame, text="Đóng",
            command=self.dialog.destroy, width=15
        ).pack(side=tk.RIGHT)

        self.dialog.protocol("WM_DELETE_WINDOW", self.dialog.destroy)
        self.dialog.update()

    def _retry(self) -> None:
        self.error_label.pack_forget()
        for widget in self.btn_frame.winfo_children():
            widget.destroy()
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        if self.retry_callback:
            self.retry_callback()
```

- [ ] **Step 2: Commit**

```bash
git add ui/pdf_import_dialog.py
git commit -m "feat: add import progress dialog with step tracking and retry"
```

---

### Task 8: Tích hợp vào Main UI

**Files:**
- Modify: `ui/excel_realtime_controller.py`
- Modify: `data/template_configs/dialog_config.json`

- [ ] **Step 1: Thêm import ở đầu file `ui/excel_realtime_controller.py`**

Thêm sau dòng `from ui.size_quantity_input_dialog import SizeQuantityInputDialog`:

```python
from excel_automation.pdf_po_parser import PDFPOParser
```

- [ ] **Step 2: Sửa buttons_config — ẩn "Quét Sizes", thêm style xanh lá**

Trong method `_create_widgets`, thay đổi block buttons_config (khoảng dòng 157-197):

Thêm Green style cạnh Yellow style:

```python
        style.configure('Green.TButton', background='#4CAF50')
        style.map('Green.TButton',
            background=[('active', '#66BB6A'), ('pressed', '#388E3C')])
```

Sửa `buttons_config` — bỏ dòng Quét Sizes:

```python
        buttons_config: List[Tuple[str, Callable]] = [
            ("👁️ Ẩn dòng ngay", self._hide_rows_realtime),
            ("👁️‍🗨️ Hiện tất cả", self._show_all_rows),
            ("📝 Nhập Số Lượng Size", self._input_size_quantities),
            ("💾 Ghi vào Excel", self._write_quantities_to_excel),
            ("📦 Xuất Danh Sách Thùng", self._export_box_list),
            ("📄 Đọc PDF", self._open_pdf_reader),
        ]
```

Thêm nút Import PO trước vòng loop `for text, command`, giống cách `update_po_btn` được tạo:

```python
        self.import_pdf_btn = ttk.Button(
            self.action_frame,
            text="📄 Import PO từ PDF",
            command=self._import_po_from_pdf,
            width=20,
            style='Green.TButton'
        )
        self.action_buttons.append(self.import_pdf_btn)
```

- [ ] **Step 3: Thêm method `_import_po_from_pdf`**

Thêm vào class `ExcelRealtimeController`, sau method `_update_po` (khoảng dòng 1160):

```python
    def _import_po_from_pdf(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Cảnh báo", "Vui lòng mở file Excel trước!")
            return

        file_path = filedialog.askopenfilename(
            title="Chọn file PDF Purchase Order",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not file_path:
            return

        from ui.pdf_import_dialog import ImportProgressDialog, PDFImportDialog

        progress = ImportProgressDialog(self.root)
        pdf_data = None
        available_sizes = []

        def run_parse_steps():
            nonlocal pdf_data, available_sizes

            try:
                progress.start_step(0)
                progress.complete_step(0)

                progress.start_step(1)
                pdf_data = PDFPOParser.parse(file_path)
                progress.complete_step(1)

                progress.start_step(2)
                available_sizes = self.com_manager.scan_sizes()
                self.available_sizes = available_sizes
                progress.complete_step(2)

                progress.dialog.destroy()

                PDFImportDialog(
                    self.root,
                    pdf_data,
                    available_sizes,
                    lambda po, color, sizes: self._execute_import(po, color, sizes)
                )

            except Exception as e:
                step = progress.current_step
                logger.error(f"Lỗi tại bước {step}: {e}")
                progress.show_error(step, str(e), run_parse_steps)

        run_parse_steps()

    def _execute_import(self, po: str, color: str, size_quantities: Dict[str, int]) -> None:
        from ui.pdf_import_dialog import ImportProgressDialog

        progress = ImportProgressDialog(self.root)

        for i in range(3):
            progress.complete_step(i)

        def run_write_steps(start_from: int = 3):
            try:
                if start_from <= 3:
                    progress.start_step(3)
                    po_manager = POUpdateManager(self.config)
                    po_manager.update_po_bulk(self.com_manager.worksheet, po)
                    progress.complete_step(3)

                if start_from <= 4:
                    progress.start_step(4)
                    color_manager = ColorCodeUpdateManager(self.config)
                    color_manager.update_color_code_bulk(self.com_manager.worksheet, color)
                    progress.complete_step(4)

                if start_from <= 5:
                    progress.start_step(5)
                    self._apply_imported_sizes(size_quantities)
                    progress.complete_step(5)

                progress.start_step(6)
                self._update_po_color_display()
                self._update_box_count_display()
                progress.complete_step(6)

                progress.finish()

                self.status_label.config(
                    text=f"Import thành công: PO={po}, Color={color}, {len(size_quantities)} sizes"
                )
                messagebox.showinfo(
                    "Thành Công",
                    f"Đã import PO từ PDF:\n\n"
                    f"PO: {po}\n"
                    f"Color: {color}\n"
                    f"Sizes: {len(size_quantities)}\n"
                    f"Total Qty: {sum(size_quantities.values()):,}"
                )

            except Exception as e:
                step = progress.current_step
                logger.error(f"Lỗi tại bước {step}: {e}")
                progress.show_error(step, str(e), lambda: run_write_steps(step))

        run_write_steps()

    def _apply_imported_sizes(self, size_quantities: Dict[str, int]) -> None:
        if not self.available_sizes:
            self.available_sizes = self.com_manager.scan_sizes()

        if not self.checkboxes:
            self._scan_sizes()

        for size, qty in size_quantities.items():
            if size in self.checkboxes:
                self.checkboxes[size].set(True)
                entry = self.quantity_entries.get(size)
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, str(qty))

        self._update_box_count_display()
        self._reset_auto_save_timer()
```

- [ ] **Step 4: Thêm kích thước dialog vào `data/template_configs/dialog_config.json`**

Thêm entry `"pdf_import"` vào key `"dialogs"`:

```json
"pdf_import": { "width": 550, "height": 600 }
```

- [ ] **Step 5: Chạy app kiểm tra nút hiển thị**

Run: `python excel_realtime_controller.py`
Expected: Nút "📄 Import PO từ PDF" hiển thị đầu tiên, màu xanh lá. Nút "Quét Sizes" không còn.

- [ ] **Step 6: Commit**

```bash
git add ui/excel_realtime_controller.py ui/pdf_import_dialog.py data/template_configs/dialog_config.json
git commit -m "feat: integrate PDF import into main UI with progress dialog"
```

---

### Task 9: Manual Integration Test

**Files:** Không tạo/sửa file — chỉ test thủ công

- [ ] **Step 1: Test luồng happy path**

1. Mở app: `python excel_realtime_controller.py`
2. Mở file Excel packing list
3. Bấm "📄 Import PO từ PDF"
4. Chọn file `Test.pdf`
5. Kiểm tra progress dialog hiện bước 1→3
6. Kiểm tra preview dialog hiện đúng: PO=9013330, Color=3104, 15 sizes
7. Kiểm tra sizes khớp Excel có ✅, sizes không khớp có ⚠️
8. Thử sửa 1 size không khớp → trạng thái cập nhật realtime
9. Bấm "Xác nhận & Ghi"
10. Kiểm tra progress bước 4→7 chạy đúng %
11. Kiểm tra Excel: cột A có PO, cột E có Color, checkboxes đúng, qty đúng

- [ ] **Step 2: Test lỗi file không hợp lệ**

1. Bấm "Import PO từ PDF"
2. Chọn file PDF bất kỳ không phải PO
3. Kiểm tra progress dừng ở bước lỗi, hiện ❌ + message
4. Bấm "Đóng"

- [ ] **Step 3: Test hủy**

1. Bấm "Import PO từ PDF"
2. Chọn file Test.pdf
3. Preview dialog hiện lên → bấm "Hủy"
4. Kiểm tra không có gì thay đổi trong Excel

- [ ] **Step 4: Commit final**

```bash
git add -A
git commit -m "feat: complete PDF PO import feature with preview and progress UI"
```
