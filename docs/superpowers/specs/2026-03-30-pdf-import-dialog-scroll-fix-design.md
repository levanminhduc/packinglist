# Fix Mousewheel Scroll trong PDFImportDialog

## Van de

Trong `PDFImportDialog`, bang danh sach size + quantity dung `Canvas` + `Scrollbar` de cuon. Hien tai mousewheel chi bind truc tiep len `self.canvas`, nen khi con tro chuot dang o tren cac widget con (checkbox, entry, label) ben trong bang, mousewheel khong hoat dong — phai di chuot ra vung trong cua canvas moi cuon duoc.

Day la van de co ban cua tkinter: event `<MouseWheel>` khong tu dong bubble up tu widget con len parent canvas.

## Pham vi

- Chi sua `ui/pdf_import_dialog.py`, class `PDFImportDialog`
- Khong anh huong den `SizeQuantityInputDialog` hay bat ky dialog nao khac

## Giai phap

Dung `dialog.bind_all()` thay vi `canvas.bind()` de bat mousewheel event o cap dialog. Vi `PDFImportDialog` da dung `grab_set()` nen chi co 1 dialog active tai 1 thoi diem, khong lo conflict voi dialog khac.

### Thay doi cu the

**1. `_create_size_table()` — doi bind tu canvas sang dialog:**

Hien tai:
```python
self.canvas.bind('<MouseWheel>', self._on_canvas_mousewheel)
self.canvas.bind('<Button-4>', lambda e: self._on_canvas_mousewheel_linux(e, 1))
self.canvas.bind('<Button-5>', lambda e: self._on_canvas_mousewheel_linux(e, -1))
```

Doi thanh:
```python
self.dialog.bind_all('<MouseWheel>', self._on_canvas_mousewheel)
self.dialog.bind_all('<Button-4>', lambda e: self._on_canvas_mousewheel_linux(e, 1))
self.dialog.bind_all('<Button-5>', lambda e: self._on_canvas_mousewheel_linux(e, -1))
```

**2. `_save_size_and_close()` — unbind truoc khi destroy:**

Them vao truoc `self.dialog.destroy()`:
```python
self.dialog.unbind_all('<MouseWheel>')
self.dialog.unbind_all('<Button-4>')
self.dialog.unbind_all('<Button-5>')
```

## Ket qua mong doi

Mousewheel hoat dong o moi vi tri trong dialog — du con tro dang tren checkbox, entry, label, hay bat ky widget nao khac trong bang size-quantity.
