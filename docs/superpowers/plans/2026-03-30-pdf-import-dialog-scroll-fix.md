# Plan: Fix Mousewheel Scroll trong PDFImportDialog

Spec: `docs/superpowers/specs/2026-03-30-pdf-import-dialog-scroll-fix-design.md`

## Step 1: Doi mousewheel bind tu canvas sang dialog

File: `ui/pdf_import_dialog.py`
Method: `_create_size_table()`

Doi 3 dong bind tu `self.canvas.bind(...)` sang `self.dialog.bind_all(...)`:

```python
# Tu:
self.canvas.bind('<MouseWheel>', self._on_canvas_mousewheel)
self.canvas.bind('<Button-4>', lambda e: self._on_canvas_mousewheel_linux(e, 1))
self.canvas.bind('<Button-5>', lambda e: self._on_canvas_mousewheel_linux(e, -1))

# Thanh:
self.dialog.bind_all('<MouseWheel>', self._on_canvas_mousewheel)
self.dialog.bind_all('<Button-4>', lambda e: self._on_canvas_mousewheel_linux(e, 1))
self.dialog.bind_all('<Button-5>', lambda e: self._on_canvas_mousewheel_linux(e, -1))
```

## Step 2: Unbind truoc khi destroy dialog

File: `ui/pdf_import_dialog.py`
Method: `_save_size_and_close()`

Them 3 dong unbind truoc `self.dialog.destroy()`:

```python
def _save_size_and_close(self) -> None:
    try:
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = self.dialog.winfo_x()
        y = self.dialog.winfo_y()
        self.dialog_config.save_dialog_geometry('pdf_import', width, height, x, y)
    except Exception as e:
        logger.error(f"Loi khi luu geometry dialog: {e}")
    self.dialog.unbind_all('<MouseWheel>')
    self.dialog.unbind_all('<Button-4>')
    self.dialog.unbind_all('<Button-5>')
    self.dialog.destroy()
```

## Kiem tra

- Mo app, import PO tu PDF
- Di chuot len tren checkbox/entry/label trong bang size
- Cuon mousewheel -> bang phai cuon binh thuong
- Di chuot ra vung trong canvas -> van cuon binh thuong
- Dong dialog -> khong loi gi
