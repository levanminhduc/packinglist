# Clear Data Preserve Tot QTY — Design Spec

## Vấn đề

Khi copy sheet, `clear_quantity_columns()` xóa range cố định G:AM (col 7→39). Cột "Tot QTY" chứa công thức SUM và vị trí của nó thay đổi tùy theo số lượng sizes trong sheet:

- Sheet ít sizes: Tot QTY ở cột **N** (col 14) → nằm trong range xóa → **formula bị mất**
- Sheet nhiều sizes: Tot QTY ở cột **AN** (col 40) → nằm ngoài range → an toàn

## Yêu cầu

- Xóa tất cả giá trị số lượng trong các cột size (từ G trở đi)
- Giữ nguyên cột Tot QTY (chứa công thức SUM) và mọi cột sau nó (N.W. tot, G.W. tot)
- Tự động detect vị trí Tot QTY, không hardcode

## Solution: Dynamic Detect Tot QTY Column

### 1. Method `_detect_tot_qty_column()` — Thêm mới vào ExcelCOMManager

```
_detect_tot_qty_column() -> Optional[int]
```

Logic:
1. Quét row 14 → 18 (5 hàng header, cover mọi template variant)
2. Trong mỗi row, đọc toàn bộ range `A{row}:AZ{row}` bằng 1 COM call (bulk read)
3. Tìm cell chứa text "Tot QTY" hoặc "Total QTY" (case-insensitive)
4. Trả về column number (ví dụ: N=14, AN=40)
5. Không tìm thấy → trả về `None`

### 2. Sửa `clear_quantity_columns()` — Dùng dynamic end_col

Signature thay đổi:

```
clear_quantity_columns(start_row, end_row, start_col=7, end_col=None)
```

`end_col` default đổi từ `39` → `None`.

Logic mới:
1. Nếu `end_col` được truyền vào → dùng giá trị đó (caller override)
2. Nếu `end_col is None` → gọi `_detect_tot_qty_column()`
   - Tìm thấy Tot QTY ở col X → `end_col = X - 1`
   - Không tìm thấy → fallback `end_col = 39` (backward-compatible)
3. ClearContents range như hiện tại

Ví dụ:
- Tot QTY ở col N (14) → xóa G:M (col 7→13)
- Tot QTY ở col AN (40) → xóa G:AM (col 7→39)
- Không tìm thấy → xóa G:AM (col 7→39)

## Files thay đổi

| File | Loại | Thay đổi |
|---|---|---|
| `excel_automation/excel_com_manager.py` | Sửa | Thêm `_detect_tot_qty_column()`, sửa `clear_quantity_columns()` |
| `tests/test_excel_com_manager.py` | Sửa | Thêm tests cho detect + update tests clear |

## Không thay đổi

- `ui/excel_realtime_controller.py` — vẫn gọi `clear_quantity_columns()` như cũ
- `ui/copy_sheet_progress_dialog.py` — không ảnh hưởng
- `box_list_export_manager.py` — có logic detect riêng, giữ nguyên

## Test Cases

1. `_detect_tot_qty_column()` — tìm thấy "Tot QTY" ở row 16, trả về đúng column number
2. `_detect_tot_qty_column()` — tìm thấy "Total QTY" (variant), trả về đúng
3. `_detect_tot_qty_column()` — không tìm thấy → trả về `None`
4. `clear_quantity_columns()` — detect Tot QTY col 14 → range G:M (col 7→13)
5. `clear_quantity_columns()` — detect trả None → fallback G:AM (col 7→39)
6. `clear_quantity_columns()` — caller truyền `end_col=39` → dùng giá trị đó, skip detect
