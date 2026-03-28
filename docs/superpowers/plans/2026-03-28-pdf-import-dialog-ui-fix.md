# Fix PDFImportDialog UI Consistency — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix 4 UI inconsistencies in PDFImportDialog and add window position persistence.

**Architecture:** Two files changed — `DialogConfigManager` gets new geometry methods + DEFAULT_CONFIG entry, `PDFImportDialog` gets font fix, mousewheel binding, and geometry save/restore.

**Tech Stack:** Python, tkinter, ttk

---

### Task 1: Add `pdf_import` to DEFAULT_CONFIG + new geometry methods in DialogConfigManager

**Files:**
- Modify: `excel_automation/dialog_config_manager.py:12-45` (DEFAULT_CONFIG) and append new methods

- [ ] **Step 1: Add `pdf_import` entry to DEFAULT_CONFIG**

In `excel_automation/dialog_config_manager.py`, add `pdf_import` to the `dialogs` dict inside `DEFAULT_CONFIG`, after `pdf_reader`:

```python
            "pdf_reader": {
                "width": 700,
                "height": 500
            },
            "pdf_import": {
                "width": 600,
                "height": 550
            }
```

- [ ] **Step 2: Add `get_dialog_geometry()` method**

Add after `save_dialog_size()` (after line 99):

```python
    def get_dialog_geometry(self, dialog_name: str) -> Tuple[int, int, Optional[int], Optional[int]]:
        try:
            dialog_config = self.config.get('dialogs', {}).get(dialog_name, {})
            width = dialog_config.get('width', 400)
            height = dialog_config.get('height', 300)
            x = dialog_config.get('x')
            y = dialog_config.get('y')
            return width, height, x, y
        except Exception as e:
            logger.error(f"Lỗi khi đọc geometry dialog {dialog_name}: {e}")
            return 400, 300, None, None
```

- [ ] **Step 3: Add `save_dialog_geometry()` method**

Add right after `get_dialog_geometry()`:

```python
    def save_dialog_geometry(self, dialog_name: str, width: int, height: int, x: int, y: int) -> None:
        try:
            if 'dialogs' not in self.config:
                self.config['dialogs'] = {}

            if dialog_name not in self.config['dialogs']:
                self.config['dialogs'][dialog_name] = {}

            self.config['dialogs'][dialog_name]['width'] = width
            self.config['dialogs'][dialog_name]['height'] = height
            self.config['dialogs'][dialog_name]['x'] = x
            self.config['dialogs'][dialog_name]['y'] = y

            self._save_config()
            logger.info(f"Đã lưu geometry dialog {dialog_name}: {width}x{height}+{x}+{y}")
        except Exception as e:
            logger.error(f"Lỗi khi lưu geometry dialog {dialog_name}: {e}")
```

- [ ] **Step 4: Commit**

```bash
git add excel_automation/dialog_config_manager.py
git commit -m "feat(dialog-config): add pdf_import default + geometry save/restore methods"
```

---

### Task 2: Fix font header + add mousewheel + save/restore geometry in PDFImportDialog

**Files:**
- Modify: `ui/pdf_import_dialog.py:12-134` (PDFImportDialog class) and `ui/pdf_import_dialog.py:219-226` (_save_size_and_close)

- [ ] **Step 1: Fix font header in `_create_size_table()`**

In `ui/pdf_import_dialog.py`, replace 4 header labels (lines 101-104) from `("", 9, "bold")` to `('Arial', 9, 'bold')`:

```python
        ttk.Label(header, text="☑", width=3, font=('Arial', 9, 'bold')).pack(side=tk.LEFT)
        ttk.Label(header, text="Size", width=10, font=('Arial', 9, 'bold')).pack(side=tk.LEFT)
        ttk.Label(header, text="Qty", width=8, font=('Arial', 9, 'bold'), anchor=tk.E).pack(side=tk.LEFT)
        ttk.Label(header, text="Trạng thái", width=16, font=('Arial', 9, 'bold')).pack(side=tk.LEFT, padx=(10, 0))
```

- [ ] **Step 2: Add mousewheel methods**

Add two new methods to `PDFImportDialog` class, after `_on_check_changed()` (after line 178):

```python
    def _on_canvas_mousewheel(self, event) -> None:
        try:
            if event.delta > 0:
                self.canvas.yview_scroll(-1, "units")
            elif event.delta < 0:
                self.canvas.yview_scroll(1, "units")
        except Exception as e:
            logger.error(f"Lỗi khi xử lý canvas mouse wheel: {e}")

    def _on_canvas_mousewheel_linux(self, event, direction: int) -> None:
        try:
            self.canvas.yview_scroll(-direction, "units")
        except Exception as e:
            logger.error(f"Lỗi khi xử lý canvas mouse wheel Linux: {e}")
```

- [ ] **Step 3: Store canvas as `self.canvas` and bind mousewheel events**

In `_create_size_table()`, change local `canvas` to `self.canvas` (line 91):

```python
        self.canvas = tk.Canvas(size_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(size_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        scrollable = ttk.Frame(self.canvas)

        scrollable.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=scrollable, anchor=tk.NW)
        self.canvas.configure(yscrollcommand=scrollbar.set)
```

Then after `scrollbar.pack(side=tk.RIGHT, fill=tk.Y)` (line 133), add mousewheel bindings:

```python
        self.canvas.bind('<MouseWheel>', self._on_canvas_mousewheel)
        self.canvas.bind('<Button-4>', lambda e: self._on_canvas_mousewheel_linux(e, 1))
        self.canvas.bind('<Button-5>', lambda e: self._on_canvas_mousewheel_linux(e, -1))
```

- [ ] **Step 4: Update `__init__` to restore geometry with position**

Replace lines 34-42 in `__init__()`:

```python
        width, height = self.dialog_config.get_dialog_size('pdf_import')
        self.dialog.geometry(f"{width}x{height}")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_closing)

        self._create_widgets()
        self._center_window()
```

With:

```python
        width, height, x, y = self.dialog_config.get_dialog_geometry('pdf_import')
        if x is not None and y is not None:
            self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        else:
            self.dialog.geometry(f"{width}x{height}")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_closing)

        self._create_widgets()
        if x is None or y is None:
            self._center_window()
```

- [ ] **Step 5: Update `_save_size_and_close()` to save full geometry**

Replace lines 219-226:

```python
    def _save_size_and_close(self) -> None:
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            self.dialog_config.save_dialog_size('pdf_import', width, height)
        except Exception as e:
            logger.error(f"Lỗi khi lưu kích thước dialog: {e}")
        self.dialog.destroy()
```

With:

```python
    def _save_size_and_close(self) -> None:
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            x = self.dialog.winfo_x()
            y = self.dialog.winfo_y()
            self.dialog_config.save_dialog_geometry('pdf_import', width, height, x, y)
        except Exception as e:
            logger.error(f"Lỗi khi lưu geometry dialog: {e}")
        self.dialog.destroy()
```

- [ ] **Step 6: Commit**

```bash
git add ui/pdf_import_dialog.py
git commit -m "fix(pdf-import): fix font, add mousewheel scroll, save window position"
```
