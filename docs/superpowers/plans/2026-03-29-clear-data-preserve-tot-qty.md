# Clear Data Preserve Tot QTY — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Khi copy sheet, tự động detect vị trí cột "Tot QTY" và chỉ xóa data đến trước nó, bảo toàn công thức SUM.

**Architecture:** Thêm `_detect_tot_qty_column()` quét header rows 14→18 tìm text "Tot QTY", trả về column number. `clear_quantity_columns()` dùng kết quả này làm `end_col` thay vì hardcode 39.

**Tech Stack:** Python, win32com.client (COM), unittest/mock

**Spec:** `docs/superpowers/specs/2026-03-29-clear-data-preserve-tot-qty-design.md`

---

## File Structure

| File | Loại | Trách nhiệm |
|---|---|---|
| `excel_automation/excel_com_manager.py` | Sửa | Thêm `_detect_tot_qty_column()`, sửa `clear_quantity_columns()` default `end_col=None` |
| `tests/test_excel_com_manager.py` | Sửa | Thêm `TestDetectTotQtyColumn`, update `TestClearQuantityColumnsBulk` |

---

### Task 1: Thêm `_detect_tot_qty_column()` với TDD

**Files:**
- Modify: `tests/test_excel_com_manager.py`
- Modify: `excel_automation/excel_com_manager.py`

- [ ] **Step 1: Viết failing tests cho `_detect_tot_qty_column()`**

Thêm class `TestDetectTotQtyColumn` vào cuối file `tests/test_excel_com_manager.py`, TRƯỚC dòng `if __name__ == "__main__":`:

```python
class TestDetectTotQtyColumn(unittest.TestCase):

    def setUp(self):
        with patch.object(ExcelCOMManager, '__init__', lambda self, *a, **kw: None):
            self.manager = ExcelCOMManager()
            self.manager.worksheet = MagicMock()

    def _make_row_values(self, col_count: int, tot_qty_col: int = None, text: str = "Tot QTY"):
        row = [None] * col_count
        if tot_qty_col is not None:
            row[tot_qty_col - 1] = text
        return tuple(row)

    def test_finds_tot_qty_at_column_14(self):
        row_data = self._make_row_values(52, tot_qty_col=14)
        self.manager.worksheet.Range.return_value.Value = (row_data,)
        result = self.manager._detect_tot_qty_column()
        self.assertEqual(result, 14)

    def test_finds_tot_qty_at_column_40(self):
        row_data = self._make_row_values(52, tot_qty_col=40)
        self.manager.worksheet.Range.return_value.Value = (row_data,)
        result = self.manager._detect_tot_qty_column()
        self.assertEqual(result, 40)

    def test_finds_total_qty_variant(self):
        row_data = self._make_row_values(52, tot_qty_col=20, text="Total QTY")
        self.manager.worksheet.Range.return_value.Value = (row_data,)
        result = self.manager._detect_tot_qty_column()
        self.assertEqual(result, 20)

    def test_case_insensitive(self):
        row_data = self._make_row_values(52, tot_qty_col=14, text="tot qty")
        self.manager.worksheet.Range.return_value.Value = (row_data,)
        result = self.manager._detect_tot_qty_column()
        self.assertEqual(result, 14)

    def test_returns_none_when_not_found(self):
        row_data = self._make_row_values(52)
        self.manager.worksheet.Range.return_value.Value = (row_data,)
        result = self.manager._detect_tot_qty_column()
        self.assertIsNone(result)

    def test_returns_none_when_no_worksheet(self):
        self.manager.worksheet = None
        result = self.manager._detect_tot_qty_column()
        self.assertIsNone(result)
```

- [ ] **Step 2: Run test — verify FAIL**

Run: `pytest tests/test_excel_com_manager.py::TestDetectTotQtyColumn -v`
Expected: FAIL — `AttributeError: 'ExcelCOMManager' object has no attribute '_detect_tot_qty_column'`

- [ ] **Step 3: Implement `_detect_tot_qty_column()`**

Thêm method sau vào `excel_automation/excel_com_manager.py`, ngay TRƯỚC method `clear_quantity_columns()` (trước dòng `def clear_quantity_columns`):

```python
    def _detect_tot_qty_column(self) -> Optional[int]:
        if self.worksheet is None:
            return None

        try:
            for row in range(14, 19):
                range_str = f"A{row}:AZ{row}"
                row_values = self.worksheet.Range(range_str).Value

                if row_values is None:
                    continue

                cells = row_values[0] if isinstance(row_values[0], tuple) else row_values

                for col_idx, cell_value in enumerate(cells):
                    if cell_value is not None and isinstance(cell_value, str):
                        cell_lower = cell_value.strip().lower()
                        if "tot qty" in cell_lower:
                            col_number = col_idx + 1
                            logger.info(f"Tìm thấy Tot QTY tại row {row}, col {col_number}")
                            return col_number

            logger.warning("Không tìm thấy cột Tot QTY trong row 14-18")
            return None

        except Exception as e:
            logger.warning(f"Lỗi khi detect cột Tot QTY: {e}")
            return None
```

- [ ] **Step 4: Run test — verify PASS**

Run: `pytest tests/test_excel_com_manager.py::TestDetectTotQtyColumn -v`
Expected: ALL 6 tests PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_excel_com_manager.py excel_automation/excel_com_manager.py
git commit -m "feat: add _detect_tot_qty_column() to ExcelCOMManager"
```

---

### Task 2: Sửa `clear_quantity_columns()` dùng dynamic end_col

**Files:**
- Modify: `tests/test_excel_com_manager.py`
- Modify: `excel_automation/excel_com_manager.py`

- [ ] **Step 1: Thêm tests mới cho dynamic end_col behavior**

Thêm 3 test methods vào class `TestClearQuantityColumnsBulk` trong `tests/test_excel_com_manager.py`, sau method `test_raises_if_no_worksheet`:

```python
    def test_uses_detected_tot_qty_column_as_end(self):
        with patch.object(self.manager, '_detect_tot_qty_column', return_value=14):
            self.manager.clear_quantity_columns(start_row=19, end_row=59)

        self.manager.worksheet.Range.assert_called_once_with("G19:M59")
        self.manager.worksheet.Range.return_value.ClearContents.assert_called_once()

    def test_fallback_to_39_when_tot_qty_not_found(self):
        with patch.object(self.manager, '_detect_tot_qty_column', return_value=None):
            self.manager.clear_quantity_columns(start_row=19, end_row=59)

        self.manager.worksheet.Range.assert_called_once_with("G19:AM59")

    def test_explicit_end_col_skips_detect(self):
        with patch.object(self.manager, '_detect_tot_qty_column') as mock_detect:
            self.manager.clear_quantity_columns(start_row=19, end_row=59, end_col=25)

        mock_detect.assert_not_called()
        self.manager.worksheet.Range.assert_called_once_with("G19:Y59")
```

- [ ] **Step 2: Run test — verify FAIL cho tests mới**

Run: `pytest tests/test_excel_com_manager.py::TestClearQuantityColumnsBulk::test_uses_detected_tot_qty_column_as_end tests/test_excel_com_manager.py::TestClearQuantityColumnsBulk::test_fallback_to_39_when_tot_qty_not_found tests/test_excel_com_manager.py::TestClearQuantityColumnsBulk::test_explicit_end_col_skips_detect -v`
Expected: FAIL — các test mới gọi không truyền `end_col` nhưng code hiện tại default `end_col=39` không gọi detect

- [ ] **Step 3: Sửa `clear_quantity_columns()` signature và logic**

Thay toàn bộ method `clear_quantity_columns` trong `excel_automation/excel_com_manager.py`:

```python
    def clear_quantity_columns(self, start_row: Optional[int] = None,
                               end_row: Optional[int] = None,
                               start_col: int = 7,
                               end_col: Optional[int] = None) -> int:
        if self.worksheet is None:
            raise RuntimeError("Chưa chọn worksheet nào")

        start_row = start_row or self.config.get_start_row()
        end_row = end_row or self.detect_end_row()

        if end_col is None:
            detected = self._detect_tot_qty_column()
            end_col = (detected - 1) if detected else 39

        try:
            if self.excel_app:
                self.excel_app.ScreenUpdating = False

            start_col_letter = self._number_to_column_letter(start_col)
            end_col_letter = self._number_to_column_letter(end_col)
            range_str = f"{start_col_letter}{start_row}:{end_col_letter}{end_row}"

            target_range = self.worksheet.Range(range_str)
            cleared_count = int(self.excel_app.WorksheetFunction.CountA(target_range))
            target_range.ClearContents()

            if self.excel_app:
                self.excel_app.ScreenUpdating = True

            logger.info(f"Đã xóa {cleared_count} ô số lượng ({range_str})")
            return cleared_count

        except Exception as e:
            if self.excel_app:
                self.excel_app.ScreenUpdating = True
            logger.error(f"Lỗi khi xóa số lượng: {e}")
            raise RuntimeError(f"Không thể xóa số lượng: {str(e)}")
```

- [ ] **Step 4: Run ALL tests — verify PASS**

Run: `pytest tests/test_excel_com_manager.py -v`
Expected: ALL tests PASS (cả tests cũ lẫn mới — tests cũ truyền `end_col=39` explicit nên không bị ảnh hưởng)

- [ ] **Step 5: Commit**

```bash
git add tests/test_excel_com_manager.py excel_automation/excel_com_manager.py
git commit -m "feat: clear_quantity_columns() auto-detect Tot QTY column to preserve formulas"
```
