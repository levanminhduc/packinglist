# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel Automation for textile/garment packing lists. Windows-only Python app using tkinter GUI and COM automation to control live Microsoft Excel instances. Domain: apparel export — PO numbers, size ranges (numeric 044–060 or letter XS–XXXL), color codes, carton allocation.

Three entry points:
- `excel_realtime_controller.py` — Main GUI app (tkinter + COM)
- `excel_viewer.py` — Read-only Excel viewer with validation
- `main.py` — CLI interactive menu for batch scripts

## Commands

```bash
# Setup
python -m venv venv && venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env

# Run apps
python excel_realtime_controller.py   # Main GUI
python excel_viewer.py                # Viewer
python main.py                        # CLI menu

# Tests
pytest tests/                              # All tests
pytest tests/ --cov=excel_automation       # With coverage
pytest tests/test_reader.py -v             # Single test file
pytest tests/test_size_filter.py -v        # Single test file

# Build to .exe
pyinstaller excel_realtime_controller.spec --clean --noconfirm
# Output: dist/ExcelRealtimeController/ExcelRealtimeController.exe
```

## Architecture

```
excel_automation/   # Core business logic — NO tkinter/UI imports
ui/                 # All tkinter GUI code
config/             # Settings from .env via python-dotenv
scripts/            # Standalone automation scripts
tests/              # pytest + unittest
data/
  template_configs/ # JSON configs (size_filter, dialog, box_list)
  validation_rules/ # JSON validation rule definitions
  templates/        # Excel template files
```

**Business logic and UI are strictly separated.** `excel_automation/` contains zero GUI imports; `ui/` handles all tkinter code.

### Key Patterns

- **Manager pattern**: Each domain concern has a dedicated manager class (`ExcelCOMManager`, `SizeFilterManager`, `POUpdateManager`, `ColorCodeUpdateManager`, `SizeQuantityDisplayManager`, `BoxListExportManager`)
- **JSON config persistence**: Each manager has its own config class with `DEFAULT_CONFIG`, `_load_config()`, `_save_config()`, `_merge_with_defaults()`. Configs live in `data/template_configs/`
- **Dual-mode paths**: `excel_automation/path_helper.py` detects `sys.frozen` — script mode uses project root, `.exe` mode uses `%APPDATA%/ExcelRealtimeController` for user data
- **COM automation**: `excel_com_manager.py` uses `win32com.client` to control a live Excel instance. Bulk operations wrap with `ScreenUpdating = False/True` for performance
- **Size normalization**: `utils.normalize_size_value()` handles COM float returns (38.0 -> "038"), rounds half-sizes, preserves letter sizes

### Key Constants

- Data rows start at row 19 (rows 1–18 are header/info in packing list template)
- Default size column: F, default row range: 19–59 (configurable via `size_filter_config.json`)
- Auto-save delay: 10 seconds (`AUTO_SAVE_DELAY_MS = 10000`)
- Auto-refresh polling: 3 seconds

### Main UI Module

`ui/excel_realtime_controller.py` (~63KB) contains the primary `ExcelRealtimeController` class with all main window logic: file open, sheet selection, size scanning, checkbox controls, PO/color/quantity dialogs, auto-save, auto-refresh.

## Agent & Subagent Rules

- **Codebase discovery**: Khi cần tìm hiểu codebase (tìm file, hiểu kiến trúc, tìm symbol, class, function), luôn ưu tiên dùng `mcp__auggie__codebase-retrieval` (Context Engine) thay vì đọc từng file thủ công. Áp dụng cho cả main agent và tất cả subagent (superpowers, Plan, Explore, general-purpose).

## Conventions

- **Language**: Always reply in Vietnamese. All user-facing text and error messages are in Vietnamese
- **No code comments**: Do not add comments in code
- **No markdown documentation**: Do not create markdown documentation files
- **Type hints**: Used throughout all modules
- **Logging**: Every module uses `logging.getLogger(__name__)`; logs go to `logs/` directory
- **Error handling**: Public methods use try/except with `logger.error()` and re-raise as `RuntimeError` with Vietnamese messages
- **Callback pattern**: UI dialogs accept `on_save_callback: Callable` parameters

## Dependencies

Key libraries: `pandas`, `openpyxl`, `xlsxwriter`, `pywin32` (COM), `pdfplumber`, `pytesseract`, `Pillow`. Full list in `requirements.txt`.

Requires Windows with Microsoft Excel installed for COM automation features.
