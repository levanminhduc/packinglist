# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from pathlib import Path

block_cipher = None

project_root = os.path.abspath('.')

a = Analysis(
    ['excel_realtime_controller.py'],
    pathex=[project_root],
    binaries=[],
    datas=[
        ('config', 'config'),
        ('data/template_configs', 'data/template_configs'),
        ('data/validation_rules', 'data/validation_rules'),
        ('data/templates', 'data/templates'),
        ('ui', 'ui'),
        ('excel_automation', 'excel_automation'),
    ],
    hiddenimports=[
        'win32com',
        'win32com.client',
        'win32api',
        'pythoncom',
        'pywintypes',
        'pandas',
        'openpyxl',
        'xlsxwriter',
        'xlrd',
        'xlwt',
        'xlutils',
        'dotenv',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'config',
        'config.settings',
        'ui.excel_realtime_controller',
        'ui.size_filter_config_dialog',
        'ui.size_filter_dialog',
        'ui.settings_dialog',
        'ui.po_update_dialog',
        'ui.color_code_update_dialog',
        'ui.size_quantity_input_dialog',
        'ui.excel_viewer_window',
        'ui.ui_config',
        'excel_automation.excel_com_manager',
        'excel_automation.size_filter_config',
        'excel_automation.size_filter',
        'excel_automation.dialog_config_manager',
        'excel_automation.size_quantity_display_manager',
        'excel_automation.box_list_export_config',
        'excel_automation.box_list_export_manager',
        'excel_automation.po_update_manager',
        'excel_automation.color_code_update_manager',
        'excel_automation.carton_allocation_calculator',
        'excel_automation.formatter',
        'excel_automation.processor',
        'excel_automation.reader',
        'excel_automation.writer',
        'excel_automation.utils',
        'excel_automation.validator',
        'excel_automation.validation_rules',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ExcelRealtimeController',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ExcelRealtimeController',
)

