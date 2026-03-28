import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Optional, Callable, Tuple, TypeVar, Any
from pathlib import Path
import logging
import math
import re
import time

from excel_automation.excel_com_manager import ExcelCOMManager
from excel_automation.size_filter_config import SizeFilterConfig
from excel_automation.dialog_config_manager import DialogConfigManager
from excel_automation.size_quantity_display_manager import SizeQuantityDisplayManager
from excel_automation.box_list_export_config import BoxListExportConfig
from excel_automation.box_list_export_manager import BoxListExportManager
from excel_automation.carton_allocation_calculator import (
    CartonAllocationCalculator,
    AllocationResult
)
from excel_automation.po_update_manager import POUpdateManager
from excel_automation.color_code_update_manager import ColorCodeUpdateManager
from ui.size_quantity_input_dialog import SizeQuantityInputDialog
from excel_automation.pdf_po_parser import PDFPOParser

logger = logging.getLogger(__name__)


class ExcelRealtimeController:

    AUTO_SAVE_DELAY_MS = 10000

    def __init__(self, root: tk.Tk):
        self.root = root
        self.config = SizeFilterConfig()
        self.dialog_config = DialogConfigManager()
        self.com_manager: Optional[ExcelCOMManager] = None
        self.current_file: Optional[str] = None
        self.sheet_names: List[str] = []
        self.current_sheet: Optional[str] = None
        self.available_sizes: List[str] = []
        self.checkboxes: Dict[str, tk.BooleanVar] = {}
        self.action_buttons: List[ttk.Button] = []
        self.action_frame: Optional[ttk.Frame] = None

        self.quantity_entries: Dict[str, ttk.Entry] = {}
        self.items_per_box: Optional[int] = None
        self.allocation_result: Optional[AllocationResult] = None
        self.box_count_frame: Optional[ttk.LabelFrame] = None
        self.total_qty_label: Optional[ttk.Label] = None
        self.box_count_label: Optional[ttk.Label] = None
        self.allocation_text: Optional[tk.Text] = None

        self._auto_save_timer_id: Optional[str] = None
        self._auto_save_pending: bool = False

        self._auto_refresh_sizes_timer_id: Optional[str] = None
        self._auto_refresh_interval: int = 3000
        self._cached_sizes: List[str] = []

        self.po_updated: bool = False
        self.color_updated: bool = False
        self.update_po_btn: Optional[ttk.Button] = None
        self.update_color_btn: Optional[ttk.Button] = None
        self.current_po_label: Optional[ttk.Label] = None
        self.current_color_label: Optional[ttk.Label] = None
        self.current_mahang_label: Optional[ttk.Label] = None

        self._setup_window()
        self._create_widgets()
    
    def _setup_window(self) -> None:
        self.root.title("Nhập Packing List - by Chồng Thi")

        width, height, x, y = self.dialog_config.get_main_window_geometry()
        if x is not None and y is not None:
            self.root.geometry(f"{width}x{height}+{x}+{y}")
        else:
            self.root.geometry(f"{width}x{height}")

        self.root.resizable(True, True)

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _create_widgets(self) -> None:
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        file_frame = ttk.LabelFrame(main_frame, text="File Excel", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(
            file_frame,
            text="📂 Chọn File Excel",
            command=self._open_file,
            width=20
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="Chưa mở file nào", foreground="gray")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        sheet_frame = ttk.LabelFrame(main_frame, text="Chọn Sheet", padding=10)
        sheet_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(sheet_frame, text="Sheet:").pack(side=tk.LEFT, padx=(0, 5))

        self.sheet_combobox = ttk.Combobox(sheet_frame, state="readonly", width=30)
        self.sheet_combobox.pack(side=tk.LEFT, padx=(0, 10))
        self.sheet_combobox.bind('<<ComboboxSelected>>', self._on_sheet_changed)

        ttk.Button(
            sheet_frame,
            text="🔄 Reload",
            command=self._reload_sheets,
            width=12
        ).pack(side=tk.LEFT, padx=(0, 10))

        self.sheet_status_label = ttk.Label(sheet_frame, text="", foreground="gray")
        self.sheet_status_label.pack(side=tk.LEFT)
        
        config_frame = ttk.LabelFrame(main_frame, text="Cấu hình Lọc", padding=10)
        config_frame.pack(fill=tk.X, pady=(0, 10))

        config_info = (
            f"Cột: {self.config.get_column()} | "
            f"Dòng: {self.config.get_start_row()}-{self.config.get_end_row()}"
        )
        self.config_info_label = ttk.Label(config_frame, text=config_info, foreground="blue")
        self.config_info_label.pack(side=tk.LEFT, anchor=tk.W)

        ttk.Button(
            config_frame,
            text="⚙️ Settings",
            command=self._open_config_settings,
            width=12
        ).pack(side=tk.RIGHT)

        info_frame = ttk.LabelFrame(main_frame, text="Thông Tin Hiện Tại", padding=10)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        info_row_frame = ttk.Frame(info_frame)
        info_row_frame.pack(fill=tk.X)

        ttk.Label(info_row_frame, text="Mã Hàng:").pack(side=tk.LEFT, padx=(0, 5))
        self.current_mahang_label = ttk.Label(info_row_frame, text="Chưa mở file", foreground="gray")
        self.current_mahang_label.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(info_row_frame, text="PO:").pack(side=tk.LEFT, padx=(0, 5))
        self.current_po_label = ttk.Label(info_row_frame, text="Chưa mở file", foreground="gray")
        self.current_po_label.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(info_row_frame, text="Màu:").pack(side=tk.LEFT, padx=(0, 5))
        self.current_color_label = ttk.Label(info_row_frame, text="Chưa mở file", foreground="gray")
        self.current_color_label.pack(side=tk.LEFT)

        self.action_frame = ttk.Frame(main_frame)
        self.action_frame.pack(fill=tk.X, pady=(0, 10))

        style = ttk.Style()
        style.configure('Yellow.TButton', background='red')
        style.map('Yellow.TButton',
            background=[('active', 'gold'), ('pressed', 'orange')])

        style.configure('Green.TButton', background='#4CAF50')
        style.map('Green.TButton',
            background=[('active', '#66BB6A'), ('pressed', '#388E3C')])

        self.import_pdf_btn = ttk.Button(
            self.action_frame,
            text="📄 Import PO từ PDF",
            command=self._import_po_from_pdf,
            width=20,
            style='Green.TButton'
        )
        self.action_buttons.append(self.import_pdf_btn)

        buttons_config: List[Tuple[str, Callable]] = [
            ("👁️ Ẩn dòng ngay", self._hide_rows_realtime),
            ("👁️‍🗨️ Hiện tất cả", self._show_all_rows),
            ("📝 Nhập Số Lượng Size", self._input_size_quantities),
            ("💾 Ghi vào Excel", self._write_quantities_to_excel),
            ("📦 Xuất Danh Sách Thùng", self._export_box_list),
            ("📄 Đọc PDF", self._open_pdf_reader),
        ]

        for text, command in buttons_config:
            btn = ttk.Button(
                self.action_frame,
                text=text,
                command=command,
                width=20
            )
            self.action_buttons.append(btn)

        self.update_po_btn = ttk.Button(
            self.action_frame,
            text="📝 Update PO",
            command=self._update_po,
            width=20,
            style='Yellow.TButton'
        )
        self.action_buttons.append(self.update_po_btn)

        self.update_color_btn = ttk.Button(
            self.action_frame,
            text="🎨 Update Color",
            command=self._update_color_code,
            width=20,
            style='Yellow.TButton'
        )
        self.action_buttons.append(self.update_color_btn)

        self.action_frame.bind("<Configure>", self._rearrange_buttons)
        self.root.after(100, lambda: self._rearrange_buttons(None))

        sizes_frame = ttk.LabelFrame(main_frame, text="Chọn Sizes để Hiển thị", padding=10)
        sizes_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        button_bar = ttk.Frame(sizes_frame)
        button_bar.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(
            button_bar,
            text="✓ Chọn tất cả",
            command=self._select_all_sizes
        ).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            button_bar,
            text="✗ Bỏ chọn tất cả",
            command=self._deselect_all_sizes
        ).pack(side=tk.LEFT)

        ttk.Button(
            button_bar,
            text="🔄 Refresh Số Liệu",
            command=self._load_quantities_from_excel
        ).pack(side=tk.LEFT, padx=(10, 0))

        self.sizes_count_label = ttk.Label(
            button_bar,
            text="Chưa quét sizes",
            foreground="gray"
        )
        self.sizes_count_label.pack(side=tk.RIGHT)

        self.sizes_canvas = tk.Canvas(sizes_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(sizes_frame, orient=tk.VERTICAL, command=self.sizes_canvas.yview)
        self.scrollable_frame = ttk.Frame(self.sizes_canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.sizes_canvas.configure(scrollregion=self.sizes_canvas.bbox("all"))
        )

        self.sizes_canvas.create_window((0, 0), window=self.scrollable_frame, anchor=tk.NW)
        self.sizes_canvas.configure(yscrollcommand=scrollbar.set)

        self.sizes_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_mousewheel(event):
            self.sizes_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.sizes_canvas.bind("<MouseWheel>", _on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", _on_mousewheel)

        self.box_count_frame = ttk.LabelFrame(
            main_frame,
            text="Thông Tin Đóng Thùng",
            padding=10
        )
        self.box_count_frame.pack(fill=tk.X, pady=(0, 10))

        self.total_qty_label = ttk.Label(
            self.box_count_frame,
            text="Tổng số lượng đã nhập: 0 cái",
            foreground='gray'
        )
        self.total_qty_label.pack(anchor=tk.W)

        self.box_count_label = ttk.Label(
            self.box_count_frame,
            text="Số thùng cần đóng: 0 thùng",
            font=('Arial', 10, 'bold'),
            foreground='gray'
        )
        self.box_count_label.pack(anchor=tk.W, pady=(5, 0))

        alloc_frame = ttk.Frame(self.box_count_frame)
        alloc_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        self.allocation_text = tk.Text(
            alloc_frame,
            height=6,
            width=50,
            font=('Consolas', 9),
            state=tk.DISABLED,
            wrap=tk.WORD
        )
        alloc_scrollbar = ttk.Scrollbar(
            alloc_frame,
            orient=tk.VERTICAL,
            command=self.allocation_text.yview
        )
        self.allocation_text.configure(yscrollcommand=alloc_scrollbar.set)
        self.allocation_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        alloc_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X)

        self.status_label = ttk.Label(
            status_frame,
            text="Sẵn sàng - Vui lòng chọn file Excel",
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_label.pack(fill=tk.X)
    
    def _open_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Chọn File Excel",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("All Files", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            self.status_label.config(text=f"Đang mở file: {Path(file_path).name}...")
            self.root.update()
            
            if self.com_manager is None:
                self.com_manager = ExcelCOMManager(self.config)
            
            self.com_manager.open_excel_file(file_path)
            self.current_file = file_path
            
            self.sheet_names = self.com_manager.get_sheet_names()
            self.sheet_combobox['values'] = self.sheet_names
            
            if self.sheet_names:
                self.current_sheet = self.com_manager.current_sheet
                self.sheet_combobox.set(self.current_sheet)
                self.sheet_status_label.config(
                    text=f"({len(self.sheet_names)} sheets)",
                    foreground="blue"
                )
            
            self.file_label.config(
                text=f"📄 {Path(file_path).name}",
                foreground="black"
            )
            
            self.status_label.config(
                text=f"Đã mở file: {Path(file_path).name} - Sheet: {self.current_sheet}"
            )
            
            self._scan_sizes()

            self._update_po_color_display()
            self._highlight_update_buttons()

            self._start_auto_refresh_sizes()
            
            logger.info(f"Đã mở file qua COM: {file_path}")
            
        except Exception as e:
            logger.error(f"Lỗi khi mở file: {e}")
            messagebox.showerror(
                "Lỗi",
                f"Không thể mở file Excel:\n\n{str(e)}\n\n"
                "Vui lòng kiểm tra:\n"
                "- File có tồn tại không\n"
                "- Excel có đang mở file này không\n"
                "- Bạn có quyền truy cập file không"
            )
            self.status_label.config(text="Lỗi khi mở file")
    
    def _reload_sheets(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Cảnh báo", "Vui lòng mở file Excel trước!")
            return

        try:
            self.status_label.config(text="Đang tải lại danh sách sheets...")
            self.root.update()

            self.sheet_names = self.com_manager.get_sheet_names()
            self.sheet_combobox['values'] = self.sheet_names

            if self.current_sheet in self.sheet_names:
                self.sheet_combobox.set(self.current_sheet)
            elif self.sheet_names:
                self.sheet_combobox.set(self.sheet_names[0])

            self.sheet_status_label.config(
                text=f"({len(self.sheet_names)} sheets)",
                foreground="blue"
            )

            self.status_label.config(text=f"Đã tải lại {len(self.sheet_names)} sheets")
            logger.info(f"Đã reload {len(self.sheet_names)} sheets")

        except Exception as e:
            logger.error(f"Lỗi khi reload sheets: {e}")
            messagebox.showerror("Lỗi", f"Không thể tải lại sheets:\n{str(e)}")
            self.status_label.config(text="Lỗi khi tải lại sheets")

    def _rearrange_buttons(self, event: Optional[tk.Event] = None) -> None:
        if not self.action_frame or not self.action_buttons:
            return

        frame_width: int = self.action_frame.winfo_width()
        if frame_width <= 1:
            return

        button_width: int = 170
        max_cols: int = max(1, frame_width // button_width)

        for btn in self.action_buttons:
            btn.grid_forget()

        for idx, btn in enumerate(self.action_buttons):
            row: int = idx // max_cols
            col: int = idx % max_cols
            btn.grid(row=row, column=col, padx=5, pady=5, sticky="ew")

        for col in range(max_cols):
            self.action_frame.columnconfigure(col, weight=1)

    def _open_config_settings(self) -> None:
        try:
            from ui.size_filter_config_dialog import SizeFilterConfigDialog

            max_row = None
            if self.com_manager and self.com_manager.worksheet:
                try:
                    max_row = self.com_manager.worksheet.UsedRange.Rows.Count
                except Exception:
                    pass

            dialog = SizeFilterConfigDialog(self.root, self.config, max_row)
            self.root.wait_window(dialog.dialog)

            config_info = (
                f"Cột: {self.config.get_column()} | "
                f"Dòng: {self.config.get_start_row()}-{self.config.get_end_row()}"
            )
            self.config_info_label.config(text=config_info)

            logger.info("Đã cập nhật cấu hình lọc")

        except Exception as e:
            logger.error(f"Lỗi khi mở settings: {e}")
            messagebox.showerror("Lỗi", f"Không thể mở cấu hình:\n{str(e)}")

    def _on_sheet_changed(self, event) -> None:
        if not self.com_manager:
            return

        selected_sheet = self.sheet_combobox.get()
        if not selected_sheet or selected_sheet == self.current_sheet:
            return

        try:
            self.status_label.config(text=f"Đang chuyển sang sheet: {selected_sheet}...")
            self.root.update()

            self.com_manager.switch_sheet(selected_sheet)
            self.current_sheet = selected_sheet

            self.status_label.config(text=f"Đã chuyển sang sheet: {selected_sheet}")

            self._scan_sizes()

            self._update_po_color_display()
            self._highlight_update_buttons()

            self._start_auto_refresh_sizes()

            logger.info(f"Đã chuyển sang sheet: {selected_sheet}")

        except Exception as e:
            logger.error(f"Lỗi khi chuyển sheet: {e}")
            messagebox.showerror("Lỗi", f"Không thể chuyển sheet:\n{str(e)}")
            self.status_label.config(text="Lỗi khi chuyển sheet")
    
    def _scan_sizes(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Cảnh báo", "Vui lòng mở file Excel trước!")
            return

        try:
            self.status_label.config(text="Đang quét sizes...")
            self.root.update()

            self.available_sizes = self.com_manager.scan_sizes()

            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            self.checkboxes.clear()
            self.quantity_entries.clear()

            self.items_per_box = self._extract_items_per_box()
            logger.info(f"Items per box: {self.items_per_box}")

            if not self.available_sizes:
                ttk.Label(
                    self.scrollable_frame,
                    text="Không tìm thấy size nào",
                    foreground="red"
                ).pack(pady=20)

                self.sizes_count_label.config(
                    text="0 sizes",
                    foreground="red"
                )
                self.status_label.config(text="Không tìm thấy size nào")
                return

            num_columns = 5
            for col in range(num_columns):
                self.scrollable_frame.columnconfigure(col, weight=1, uniform="size_col")

            for idx, size in enumerate(self.available_sizes):
                row = idx // num_columns
                col = idx % num_columns

                size_frame = ttk.Frame(self.scrollable_frame)
                size_frame.grid(row=row, column=col, sticky=tk.EW, padx=2, pady=2)

                var = tk.BooleanVar(value=False)
                self.checkboxes[size] = var

                cb = ttk.Checkbutton(
                    size_frame,
                    text=size,
                    variable=var,
                    width=8,
                    command=lambda s=size: self._on_checkbox_changed(s)
                )
                cb.pack(side=tk.LEFT)

                entry = ttk.Entry(size_frame, width=5)
                entry.pack(side=tk.LEFT, padx=(2, 0))
                entry.bind('<KeyRelease>', lambda e, s=size: self._on_quantity_changed(s, e))
                entry.bind('<MouseWheel>', lambda e: self.sizes_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
                self.quantity_entries[size] = entry

            self.sizes_count_label.config(
                text=f"Tìm thấy {len(self.available_sizes)} sizes",
                foreground="green"
            )

            detected_end = self.com_manager.detect_end_row()

            self.status_label.config(
                text=f"Đã quét {len(self.available_sizes)} sizes - "
                f"Cột {self.config.get_column()} "
                f"[{self.config.get_start_row()}:{detected_end}]"
            )
            
            self._cached_sizes = self.available_sizes.copy()

            logger.info(f"Đã quét {len(self.available_sizes)} sizes")

            self._load_quantities_from_excel()
            
        except Exception as e:
            logger.error(f"Lỗi khi quét sizes: {e}")
            messagebox.showerror("Lỗi", f"Không thể quét sizes:\n{str(e)}")
            self.status_label.config(text="Lỗi khi quét sizes")

    def _load_quantities_from_excel(self) -> None:
        if not self.com_manager or not self.available_sizes:
            return

        try:
            self.status_label.config(text="Đang đọc số liệu từ Excel...")
            self.root.update()

            display_manager = SizeQuantityDisplayManager(self.config)

            current_quantities = display_manager.get_current_quantities(
                self.com_manager.worksheet,
                self.available_sizes,
                self.config.get_column()
            )

            loaded_count = 0
            for size, quantity in current_quantities.items():
                if size in self.quantity_entries and quantity is not None:
                    entry = self.quantity_entries[size]
                    entry.delete(0, tk.END)
                    entry.insert(0, str(quantity))
                    self.checkboxes[size].set(True)
                    loaded_count += 1

            if loaded_count > 0:
                self.status_label.config(
                    text=f"Đã load {loaded_count} số liệu từ Excel"
                )
                self._update_box_count_display()
            else:
                self.status_label.config(
                    text=f"Đã quét {len(self.available_sizes)} sizes - Chưa có số liệu"
                )

            logger.info(f"Đã load {loaded_count} số liệu từ Excel")

        except Exception as e:
            logger.error(f"Lỗi khi load số liệu từ Excel: {e}")
            self.status_label.config(text="Lỗi khi đọc số liệu từ Excel")
    
    def _select_all_sizes(self) -> None:
        for var in self.checkboxes.values():
            var.set(True)
        if self.quantity_entries:
            first_entry = list(self.quantity_entries.values())[0]
            first_entry.focus_set()

    def _deselect_all_sizes(self) -> None:
        for var in self.checkboxes.values():
            var.set(False)
        for entry in self.quantity_entries.values():
            entry.delete(0, tk.END)
        self._update_box_count_display()

    def _on_quantity_changed(self, size: str, event=None) -> None:
        try:
            if size not in self.quantity_entries:
                return

            entry = self.quantity_entries[size]
            value = entry.get().strip()

            if value.isdigit() and int(value) > 0:
                self.checkboxes[size].set(True)
            else:
                self.checkboxes[size].set(False)

            self._update_box_count_display()
            self._reset_auto_save_timer()

        except Exception as e:
            logger.error(f"Lỗi khi xử lý quantity changed cho size {size}: {e}")

    def _on_checkbox_changed(self, size: str) -> None:
        try:
            if size not in self.checkboxes or size not in self.quantity_entries:
                return

            is_checked = self.checkboxes[size].get()
            entry = self.quantity_entries[size]

            if is_checked:
                entry.focus_set()
            else:
                entry.delete(0, tk.END)
                self._update_box_count_display()

        except Exception as e:
            logger.error(f"Lỗi khi xử lý checkbox changed cho size {size}: {e}")

    def _reset_auto_save_timer(self) -> None:
        if self._auto_save_timer_id is not None:
            try:
                self.root.after_cancel(self._auto_save_timer_id)
            except Exception:
                pass
            self._auto_save_timer_id = None

        self._auto_save_pending = True
        self._auto_save_timer_id = self.root.after(
            self.AUTO_SAVE_DELAY_MS,
            self._perform_auto_save
        )

    def _perform_auto_save(self) -> None:
        self._auto_save_timer_id = None

        if not self._auto_save_pending:
            return

        self._auto_save_pending = False

        if not self.com_manager:
            return

        size_quantities: Dict[str, int] = {}
        for size, entry in self.quantity_entries.items():
            value = entry.get().strip()
            if value.isdigit() and int(value) > 0:
                size_quantities[size] = int(value)

        if not size_quantities:
            return

        selected_sizes = [
            size for size, var in self.checkboxes.items()
            if var.get()
        ]

        if not selected_sizes:
            return

        try:
            self.status_label.config(text="Đang tự động lưu...")
            self.root.update()

            display_manager = SizeQuantityDisplayManager(self.config)

            current_quantities = display_manager.get_current_quantities(
                self.com_manager.worksheet,
                selected_sizes,
                self.config.get_column()
            )

            if self.allocation_result and self.items_per_box:
                written_count, columns_used = display_manager.write_allocated_quantities_to_excel(
                    self.com_manager.excel_app,
                    self.com_manager.worksheet,
                    self.allocation_result,
                    selected_sizes,
                    self.config.get_column()
                )
                self.status_label.config(
                    text=f"✓ Đã tự động lưu {written_count} cells vào Excel"
                )
                logger.info(f"Auto-save: Đã ghi {written_count} cells thành công")
            else:
                written_count = display_manager.write_quantities_to_excel(
                    self.com_manager.excel_app,
                    self.com_manager.worksheet,
                    selected_sizes,
                    size_quantities,
                    current_quantities,
                    self.config.get_column()
                )
                self.status_label.config(
                    text=f"✓ Đã tự động lưu {written_count} cells vào Excel"
                )
                logger.info(f"Auto-save: Đã ghi {written_count} cells thành công")

        except Exception as e:
            logger.error(f"Lỗi khi auto-save: {e}")
            self.status_label.config(text="Lỗi khi tự động lưu")

    def _update_box_count_display(self) -> None:
        try:
            size_quantities: Dict[str, int] = {}
            total_qty = 0

            for size, entry in self.quantity_entries.items():
                value = entry.get().strip()
                if value.isdigit() and int(value) > 0:
                    qty = int(value)
                    total_qty += qty
                    size_quantities[size] = qty

            if self.total_qty_label:
                if total_qty > 0:
                    self.total_qty_label.config(
                        text=f"Tổng số lượng đã nhập: {total_qty} cái",
                        foreground='blue'
                    )
                else:
                    self.total_qty_label.config(
                        text="Tổng số lượng đã nhập: 0 cái",
                        foreground='gray'
                    )

            if not self.items_per_box or self.items_per_box == 0:
                if self.box_count_label:
                    self.box_count_label.config(
                        text="Chưa đọc được items per box từ Excel",
                        foreground='orange'
                    )
                return

            if size_quantities:
                calc = CartonAllocationCalculator(self.items_per_box)
                self.allocation_result = calc.get_full_result(size_quantities)
                self._update_allocation_display()

                if self.box_count_label:
                    total_boxes = self.allocation_result.total_boxes
                    if total_boxes > 0:
                        self.box_count_label.config(
                            text=f"Tổng: {total_boxes} thùng "
                                 f"({self.allocation_result.total_full_boxes} nguyên + "
                                 f"{self.allocation_result.total_combined_boxes} ghép)",
                            foreground='green'
                        )
                    else:
                        self.box_count_label.config(
                            text="Số thùng cần đóng: 0 thùng",
                            foreground='gray'
                        )
            else:
                self.allocation_result = None
                if self.allocation_text:
                    self.allocation_text.config(state=tk.NORMAL)
                    self.allocation_text.delete('1.0', tk.END)
                    self.allocation_text.config(state=tk.DISABLED)
                if self.box_count_label:
                    self.box_count_label.config(
                        text="Số thùng cần đóng: 0 thùng",
                        foreground='gray'
                    )

        except Exception as e:
            logger.error(f"Lỗi khi update box count display: {e}", exc_info=True)

    def _update_allocation_display(self) -> None:
        if not self.allocation_text or not self.allocation_result:
            return

        try:
            self.allocation_text.config(state=tk.NORMAL)
            self.allocation_text.delete('1.0', tk.END)

            self.allocation_text.insert(tk.END, "=== CHI TIẾT PHÂN BỔ ===\n\n")

            for size, alloc in self.allocation_result.allocations.items():
                if alloc.remainder > 0:
                    line = f"  {size}: {alloc.total_pcs} pcs -> {alloc.full_boxes} thùng ({alloc.full_qty}) + {alloc.remainder} dư\n"
                else:
                    line = f"  {size}: {alloc.total_pcs} pcs -> {alloc.full_boxes} thùng ({alloc.full_qty})\n"
                self.allocation_text.insert(tk.END, line)

            if self.allocation_result.combined_cartons:
                self.allocation_text.insert(tk.END, "\n=== THÙNG GHÉP ===\n\n")

                for i, carton in enumerate(self.allocation_result.combined_cartons, 1):
                    details = ' + '.join([f'{s}({q})' for s, q in carton.quantities.items()])
                    full_marker = "[FULL]" if carton.is_full(self.items_per_box) else ""
                    line = f"  Thùng {i}: {details} = {carton.total_pcs} pcs {full_marker}\n"
                    self.allocation_text.insert(tk.END, line)

            self.allocation_text.config(state=tk.DISABLED)

        except Exception as e:
            logger.error(f"Lỗi khi update allocation display: {e}", exc_info=True)

    def _validate_quantities(self) -> bool:
        for size, entry in self.quantity_entries.items():
            value = entry.get().strip()

            if value == "":
                continue

            try:
                quantity = int(value)
                if quantity < 1 or quantity > 5000:
                    messagebox.showerror(
                        "Lỗi Validation",
                        f"Size {size}: Số lượng phải từ 1 đến 5000!\n\n"
                        f"Giá trị nhập: {quantity}"
                    )
                    entry.focus_set()
                    return False
            except ValueError:
                messagebox.showerror(
                    "Lỗi Validation",
                    f"Size {size}: Số lượng phải là số nguyên!\n\n"
                    f"Giá trị nhập: '{value}'"
                )
                entry.focus_set()
                return False

        return True

    def _write_quantities_to_excel(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Cảnh báo", "Vui lòng mở file Excel trước!")
            return

        selected_sizes = [
            size for size, var in self.checkboxes.items()
            if var.get()
        ]

        if not selected_sizes:
            messagebox.showwarning(
                "Cảnh báo",
                "Vui lòng chọn ít nhất một size để ghi số lượng!"
            )
            return

        if not self._validate_quantities():
            return

        size_quantities: Dict[str, int] = {}
        for size in selected_sizes:
            if size in self.quantity_entries:
                value = self.quantity_entries[size].get().strip()
                if value.isdigit() and int(value) > 0:
                    size_quantities[size] = int(value)

        if not size_quantities:
            messagebox.showwarning(
                "Cảnh báo",
                "Vui lòng nhập số lượng cho ít nhất một size!"
            )
            return

        try:
            self.status_label.config(text="Đang ghi số lượng vào Excel...")
            self.root.update()

            display_manager = SizeQuantityDisplayManager(self.config)

            current_quantities = display_manager.get_current_quantities(
                self.com_manager.worksheet,
                selected_sizes,
                self.config.get_column()
            )

            if self.allocation_result and self.items_per_box:
                written_count, columns_used = display_manager.write_allocated_quantities_to_excel(
                    self.com_manager.excel_app,
                    self.com_manager.worksheet,
                    self.allocation_result,
                    selected_sizes,
                    self.config.get_column()
                )

                result = self.allocation_result
                details_lines = []
                for size, alloc in result.allocations.items():
                    if alloc.remainder > 0:
                        details_lines.append(
                            f"  {size}: {alloc.total_pcs} pcs -> {alloc.full_boxes} thùng + {alloc.remainder} dư"
                        )
                    else:
                        details_lines.append(
                            f"  {size}: {alloc.total_pcs} pcs -> {alloc.full_boxes} thùng"
                        )

                if result.combined_cartons:
                    details_lines.append("\nThùng ghép:")
                    for i, carton in enumerate(result.combined_cartons, 1):
                        detail = ' + '.join([f'{s}({q})' for s, q in carton.quantities.items()])
                        details_lines.append(f"  Thùng {i}: {detail} = {carton.total_pcs} pcs")

                details = "\n".join(details_lines)

                messagebox.showinfo(
                    "Thành Công",
                    f"Đã ghi {written_count} cells, {columns_used} cột!\n"
                    f"Tổng: {result.total_boxes} thùng "
                    f"({result.total_full_boxes} nguyên + {result.total_combined_boxes} ghép)\n\n"
                    f"Chi tiết:\n{details}"
                )

                self.status_label.config(
                    text=f"Đã ghi {result.total_boxes} thùng ({result.total_full_boxes} nguyên + {result.total_combined_boxes} ghép)"
                )
                logger.info(f"Đã ghi {written_count} cells, {result.total_boxes} thùng thành công")

            else:
                written_count = display_manager.write_quantities_to_excel(
                    self.com_manager.excel_app,
                    self.com_manager.worksheet,
                    selected_sizes,
                    size_quantities,
                    current_quantities,
                    self.config.get_column()
                )

                details = "\n".join([
                    f"  Size {size}: {qty} pcs"
                    for size, qty in size_quantities.items()
                ])

                messagebox.showinfo(
                    "Thành Công",
                    f"Đã ghi {written_count} cells số lượng vào Excel!\n\n"
                    f"Chi tiết:\n{details}"
                )

                self.status_label.config(text=f"Đã ghi {written_count} cells số lượng")
                logger.info(f"Đã ghi {written_count} cells số lượng thành công")

        except Exception as e:
            logger.error(f"Lỗi khi ghi số lượng vào Excel: {e}", exc_info=True)
            messagebox.showerror(
                "Lỗi",
                f"Không thể ghi số lượng vào Excel:\n\n{str(e)}"
            )
            self.status_label.config(text="Lỗi khi ghi số lượng")

    def _hide_rows_realtime(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Cảnh báo", "Vui lòng mở file Excel trước!")
            return
        
        selected_sizes = [
            size for size, var in self.checkboxes.items()
            if var.get()
        ]
        
        if not selected_sizes:
            response = messagebox.askyesno(
                "Cảnh báo",
                "Bạn chưa chọn size nào!\n\n"
                "Tất cả dòng sẽ bị ẩn.\n\n"
                "Bạn có chắc muốn tiếp tục?"
            )
            if not response:
                return
        
        try:
            self.status_label.config(text="Đang ẩn dòng real-time...")
            self.root.update()
            
            hidden_count = self.com_manager.hide_rows_realtime(selected_sizes)
            
            messagebox.showinfo(
                "Thành công",
                f"Đã ẩn {hidden_count} dòng real-time!\n\n"
                f"Số sizes được chọn: {len(selected_sizes)}\n"
                f"Số dòng bị ẩn: {hidden_count}\n\n"
                "Thay đổi đã được áp dụng trực tiếp trong Excel."
            )
            
            self.status_label.config(
                text=f"Đã ẩn {hidden_count} dòng - {len(selected_sizes)} sizes được chọn"
            )
            
            logger.info(f"Đã ẩn {hidden_count} dòng real-time")
            
        except Exception as e:
            logger.error(f"Lỗi khi ẩn dòng: {e}")
            messagebox.showerror(
                "Lỗi",
                f"Không thể ẩn dòng:\n\n{str(e)}\n\n"
                "Vui lòng kiểm tra:\n"
                "- Excel có đang mở không\n"
                "- File có bị đóng không\n"
                "- Có lỗi COM automation không"
            )
            self.status_label.config(text="Lỗi khi ẩn dòng")
    
    def _show_all_rows(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Cảnh báo", "Vui lòng mở file Excel trước!")
            return
        
        if not messagebox.askyesno(
            "Xác nhận",
            "Bạn có chắc muốn hiện lại tất cả các dòng?"
        ):
            return
        
        try:
            self.status_label.config(text="Đang hiện tất cả dòng...")
            self.root.update()
            
            self.com_manager.show_all_rows()
            
            messagebox.showinfo(
                "Thành công",
                f"Đã hiện lại tất cả dòng từ {self.config.get_start_row()} "
                f"đến {self.config.get_end_row()}!"
            )
            
            self.status_label.config(text="Đã hiện tất cả dòng")
            
            logger.info("Đã hiện tất cả dòng")
            
        except Exception as e:
            logger.error(f"Lỗi khi hiện dòng: {e}")
            messagebox.showerror("Lỗi", f"Không thể hiện dòng:\n{str(e)}")
            self.status_label.config(text="Lỗi khi hiện dòng")
    
    def _update_color_code(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Cảnh báo", "Vui lòng mở file Excel trước!")
            return

        try:
            from excel_automation.color_code_update_manager import ColorCodeUpdateManager
            from ui.color_code_update_dialog import ColorCodeUpdateDialog

            color_manager = ColorCodeUpdateManager(self.config)
            current_color = color_manager.get_current_color_code(self.com_manager.worksheet)
            
            start_row, end_row = color_manager.get_data_range(self.com_manager.worksheet)

            def on_save(new_color: str) -> None:
                try:
                    self.status_label.config(text=f"Đang cập nhật mã màu thành '{new_color}'...")
                    self.root.update()

                    updated_count = color_manager.update_color_code_bulk(
                        self.com_manager.worksheet,
                        new_color
                    )

                    messagebox.showinfo(
                        "Thành Công",
                        f"Đã cập nhật {updated_count} dòng mã màu thành:\n\n'{new_color}"
                    )

                    self.status_label.config(text=f"Đã cập nhật mã màu: '{new_color}")
                    logger.info(f"Đã cập nhật {updated_count} dòng mã màu thành '{new_color}'")

                    self._update_po_color_display()
                    self._reset_color_button_highlight()

                except Exception as e:
                    logger.error(f"Lỗi khi cập nhật mã màu: {e}")
                    messagebox.showerror("Lỗi", f"Không thể cập nhật mã màu:\n{str(e)}")
                    self.status_label.config(text="Lỗi khi cập nhật mã màu")

            ColorCodeUpdateDialog(self.root, current_color, on_save, self.config, end_row)

        except Exception as e:
            logger.error(f"Lỗi khi mở dialog Update Color Code: {e}")
            messagebox.showerror("Lỗi", f"Không thể mở dialog Update Color Code:\n{str(e)}")

    def _update_po(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Cảnh báo", "Vui lòng mở file Excel trước!")
            return

        try:
            from excel_automation.po_update_manager import POUpdateManager
            from ui.po_update_dialog import POUpdateDialog

            po_manager = POUpdateManager(self.config)
            current_po = po_manager.get_current_po(self.com_manager.worksheet)
            
            start_row, end_row = po_manager.get_data_range(self.com_manager.worksheet)

            def on_save(new_po: str) -> None:
                try:
                    self.status_label.config(text=f"Đang cập nhật PO thành '{new_po}'...")
                    self.root.update()

                    updated_count = po_manager.update_po_bulk(
                        self.com_manager.worksheet,
                        new_po
                    )

                    messagebox.showinfo(
                        "Thành Công",
                        f"Đã cập nhật {updated_count} dòng PO thành:\n\n{new_po}"
                    )

                    self.status_label.config(text=f"Đã cập nhật PO: {new_po}")
                    logger.info(f"Đã cập nhật {updated_count} dòng PO thành '{new_po}'")

                    self._update_po_color_display()
                    self._reset_po_button_highlight()

                except Exception as e:
                    logger.error(f"Lỗi khi cập nhật PO: {e}")
                    messagebox.showerror("Lỗi", f"Không thể cập nhật PO:\n{str(e)}")
                    self.status_label.config(text="Lỗi khi cập nhật PO")

            POUpdateDialog(self.root, current_po, on_save, self.config, end_row)

        except Exception as e:
            logger.error(f"Lỗi khi mở dialog Update PO: {e}")
            messagebox.showerror("Lỗi", f"Không thể mở dialog Update PO:\n{str(e)}")

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

    def _input_size_quantities(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Canh bao", "Vui long mo file Excel truoc!")
            return

        selected_sizes = [
            size for size, var in self.checkboxes.items()
            if var.get()
        ]

        if not selected_sizes:
            messagebox.showwarning(
                "Canh bao",
                "Vui long chon it nhat mot size de nhap so luong!"
            )
            return

        try:
            display_manager = SizeQuantityDisplayManager(self.config)

            current_quantities = display_manager.get_current_quantities(
                self.com_manager.worksheet,
                selected_sizes,
                self.config.get_column()
            )

            dialog = SizeQuantityInputDialog(
                self.root,
                selected_sizes,
                current_quantities,
                self.com_manager.worksheet
            )
            dialog.show()

            quantities = dialog.get_quantities()

            if not quantities:
                logger.info("Nguoi dung da huy hoac khong nhap so luong nao")
                return

            self.status_label.config(text="Dang ghi so luong vao Excel...")
            self.root.update()

            allocation_result = dialog.get_allocation_result()
            items_per_box = dialog.get_items_per_box()

            if allocation_result and items_per_box:
                written_count, columns_used = display_manager.write_allocated_quantities_to_excel(
                    self.com_manager.excel_app,
                    self.com_manager.worksheet,
                    allocation_result,
                    selected_sizes,
                    self.config.get_column()
                )

                result = allocation_result
                details_lines = []
                for size, alloc in result.allocations.items():
                    if alloc.remainder > 0:
                        details_lines.append(
                            f"  {size}: {alloc.total_pcs} pcs -> {alloc.full_boxes} thung + {alloc.remainder} du"
                        )
                    else:
                        details_lines.append(
                            f"  {size}: {alloc.total_pcs} pcs -> {alloc.full_boxes} thung"
                        )

                if result.combined_cartons:
                    details_lines.append("\nThung ghep:")
                    for i, carton in enumerate(result.combined_cartons, 1):
                        detail = ' + '.join([f'{s}({q})' for s, q in carton.quantities.items()])
                        details_lines.append(f"  Thung {i}: {detail} = {carton.total_pcs} pcs")

                details = "\n".join(details_lines)

                messagebox.showinfo(
                    "Thanh Cong",
                    f"Da ghi {written_count} cells, {columns_used} cot!\n"
                    f"Tong: {result.total_boxes} thung "
                    f"({result.total_full_boxes} nguyen + {result.total_combined_boxes} ghep)\n\n"
                    f"Chi tiet:\n{details}"
                )

                self.status_label.config(
                    text=f"Da ghi {result.total_boxes} thung ({result.total_full_boxes} nguyen + {result.total_combined_boxes} ghep)"
                )
                logger.info(f"Da ghi {written_count} cells, {result.total_boxes} thung thanh cong")

            else:
                written_count = display_manager.write_quantities_to_excel(
                    self.com_manager.excel_app,
                    self.com_manager.worksheet,
                    selected_sizes,
                    quantities,
                    current_quantities,
                    self.config.get_column()
                )

                details = "\n".join([
                    f"  Size {size}: {qty if qty is not None else 'Da xoa'} pcs"
                    for size, qty in quantities.items()
                ])

                messagebox.showinfo(
                    "Thanh Cong",
                    f"Da ghi {written_count} cells so luong vao Excel!\n\n"
                    f"Chi tiet:\n{details}"
                )

                self.status_label.config(text=f"Da ghi {written_count} cells so luong")
                logger.info(f"Da ghi {written_count} cells so luong thanh cong")

        except Exception as e:
            logger.error(f"Loi khi nhap so luong size: {e}", exc_info=True)
            messagebox.showerror(
                "Loi",
                f"Khong the ghi so luong vao Excel:\n\n{str(e)}"
            )
            self.status_label.config(text="Loi khi ghi so luong")

    def _extract_items_per_box(self) -> Optional[int]:
        try:
            if not self.com_manager:
                return None
            formula = self.com_manager.worksheet.Cells(18, 7).Formula
            if not formula or not isinstance(formula, str):
                return None
            import re
            match = re.search(r'/\s*(\d+)\s*$', formula)
            if match:
                return int(match.group(1))
            return None
        except Exception as e:
            logger.warning(f"Không thể đọc items_per_box từ G18: {e}")
            return None

    def _export_box_list(self) -> None:
        if not self.com_manager:
            messagebox.showwarning("Cảnh báo", "Vui lòng mở file Excel trước!")
            return

        selected_sizes = [
            size for size, var in self.checkboxes.items()
            if var.get()
        ]

        if not selected_sizes:
            messagebox.showwarning(
                "Cảnh báo",
                "Vui lòng chọn ít nhất một size để xuất danh sách thùng!"
            )
            return

        try:
            self.status_label.config(text="Đang xuất danh sách thùng...")
            self.root.update()

            config = BoxListExportConfig()
            manager = BoxListExportManager(config)

            items_per_box = self._extract_items_per_box()

            result = manager.export_box_list(
                self.com_manager.excel_app,
                self.com_manager.workbook,
                self.com_manager.worksheet,
                selected_sizes,
                items_per_box
            )

            if result.success:
                summary = result.get_summary()

                try:
                    new_sheet = manager.create_new_sheet(
                        self.com_manager.workbook,
                        self.com_manager.worksheet
                    )

                    paste_success = manager.paste_and_format_to_excel(
                        self.com_manager.workbook,
                        self.com_manager.worksheet,
                        result.box_ranges,
                        new_sheet,
                        "A",
                        1,
                        items_per_box
                    )

                    if paste_success:
                        messagebox.showinfo(
                            "Thành Công",
                            f"{summary}\n\n"
                            f"Danh sách thùng đã được xuất vào sheet mới: {new_sheet.Name}\n"
                            f"Tất cả nội dung đã được căn giữa tự động."
                        )
                    else:
                        messagebox.showinfo(
                            "Thành Công",
                            f"{summary}\n\n"
                            f"Danh sách thùng đã được copy vào clipboard.\n"
                            f"Vui lòng paste (Ctrl+V) vào Excel."
                        )
                except Exception as paste_error:
                    logger.warning(f"Không thể paste tự động: {paste_error}")
                    messagebox.showinfo(
                        "Thành Công",
                        f"{summary}\n\n"
                        f"Danh sách thùng đã được copy vào clipboard.\n"
                        f"Vui lòng paste (Ctrl+V) vào Excel."
                    )

                self.status_label.config(text=summary)
                logger.info(f"Xuất danh sách thùng thành công: {summary}")
            else:
                messagebox.showerror(
                    "Lỗi",
                    f"Không thể xuất danh sách thùng:\n\n{result.error_message}"
                )
                self.status_label.config(text="Lỗi khi xuất danh sách thùng")
                logger.error(f"Xuất danh sách thùng thất bại: {result.error_message}")

        except Exception as e:
            logger.error(f"Lỗi khi xuất danh sách thùng: {e}", exc_info=True)
            messagebox.showerror(
                "Lỗi",
                f"Không thể xuất danh sách thùng:\n\n{str(e)}"
            )
            self.status_label.config(text="Lỗi khi xuất danh sách thùng")

    def _open_pdf_reader(self) -> None:
        """Mở dialog đọc PDF."""
        try:
            from ui.pdf_reader_dialog import PdfReaderDialog
            PdfReaderDialog(self.root)
        except ImportError as e:
            messagebox.showerror(
                "Lỗi",
                f"Không thể mở tính năng đọc PDF.\n"
                f"Hãy cài đặt thư viện: pip install pdfplumber pdf2image pytesseract Pillow\n\n"
                f"Chi tiết: {e}"
            )
        except Exception as e:
            logger.error(f"Lỗi mở PDF Reader: {e}")
            messagebox.showerror("Lỗi", f"Lỗi mở PDF Reader: {e}")

    def _update_po_color_display(self) -> None:
        if not self.com_manager or not self.com_manager.worksheet:
            return

        try:
            if self.current_mahang_label and self.current_file:
                file_name = Path(self.current_file).stem
                self.current_mahang_label.config(text=file_name, foreground="blue")

            po_manager = POUpdateManager(self.config)
            current_po = po_manager.get_current_po(self.com_manager.worksheet, 'A')

            if self.current_po_label:
                if current_po:
                    self.current_po_label.config(text=current_po, foreground="blue")
                else:
                    self.current_po_label.config(text="Chưa có", foreground="gray")

            color_manager = ColorCodeUpdateManager(self.config)
            current_color = color_manager.get_current_color_code(self.com_manager.worksheet, 'E')

            if self.current_color_label:
                if current_color:
                    self.current_color_label.config(text=current_color, foreground="blue")
                else:
                    self.current_color_label.config(text="Chưa có", foreground="gray")

        except Exception as e:
            logger.error(f"Lỗi khi cập nhật hiển thị PO/Color: {e}")

    def _highlight_update_buttons(self) -> None:
        self.po_updated = False
        self.color_updated = False
        if self.update_po_btn:
            self.update_po_btn.configure(style='Yellow.TButton')
        if self.update_color_btn:
            self.update_color_btn.configure(style='Yellow.TButton')

    def _reset_po_button_highlight(self) -> None:
        self.po_updated = True
        if self.update_po_btn:
            self.update_po_btn.configure(style='TButton')

    def _reset_color_button_highlight(self) -> None:
        self.color_updated = True
        if self.update_color_btn:
            self.update_color_btn.configure(style='TButton')

    def _column_number_to_letter(self, col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result

    def _start_auto_refresh_sizes(self) -> None:
        self._stop_auto_refresh_sizes()
        self._auto_refresh_sizes_timer_id = self.root.after(
            self._auto_refresh_interval,
            self._check_sizes_changed
        )

    def _stop_auto_refresh_sizes(self) -> None:
        if self._auto_refresh_sizes_timer_id is not None:
            try:
                self.root.after_cancel(self._auto_refresh_sizes_timer_id)
            except Exception:
                pass
            self._auto_refresh_sizes_timer_id = None

    def _check_sizes_changed(self) -> None:
        try:
            if not self.com_manager:
                return

            new_sizes = self.com_manager.scan_sizes()

            if set(new_sizes) != set(self._cached_sizes):
                current_quantities: Dict[str, str] = {}
                for size, entry in self.quantity_entries.items():
                    value = entry.get().strip()
                    if value:
                        current_quantities[size] = value

                self._cached_sizes = new_sizes.copy()
                self.available_sizes = new_sizes

                for widget in self.scrollable_frame.winfo_children():
                    widget.destroy()
                self.checkboxes.clear()
                self.quantity_entries.clear()

                if not self.available_sizes:
                    ttk.Label(
                        self.scrollable_frame,
                        text="Không tìm thấy size nào",
                        foreground="red"
                    ).pack(pady=20)

                    self.sizes_count_label.config(
                        text="0 sizes",
                        foreground="red"
                    )
                else:
                    num_columns = 5
                    for col in range(num_columns):
                        self.scrollable_frame.columnconfigure(col, weight=1, uniform="size_col")

                    for idx, size in enumerate(self.available_sizes):
                        row = idx // num_columns
                        col = idx % num_columns

                        size_frame = ttk.Frame(self.scrollable_frame)
                        size_frame.grid(row=row, column=col, sticky=tk.EW, padx=2, pady=2)

                        var = tk.BooleanVar(value=False)
                        self.checkboxes[size] = var

                        cb = ttk.Checkbutton(
                            size_frame,
                            text=size,
                            variable=var,
                            width=8,
                            command=lambda s=size: self._on_checkbox_changed(s)
                        )
                        cb.pack(side=tk.LEFT)

                        entry = ttk.Entry(size_frame, width=5)
                        entry.pack(side=tk.LEFT, padx=(2, 0))
                        entry.bind('<KeyRelease>', lambda e, s=size: self._on_quantity_changed(s, e))
                        entry.bind('<MouseWheel>', lambda e: self.sizes_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
                        self.quantity_entries[size] = entry

                        if size in current_quantities:
                            entry.insert(0, current_quantities[size])
                            if current_quantities[size].isdigit() and int(current_quantities[size]) > 0:
                                var.set(True)

                    self.sizes_count_label.config(
                        text=f"Tìm thấy {len(self.available_sizes)} sizes",
                        foreground="green"
                    )

                self._update_box_count_display()

        except Exception as e:
            logger.error(f"Lỗi khi check sizes changed: {e}")
        finally:
            self._auto_refresh_sizes_timer_id = self.root.after(
                self._auto_refresh_interval,
                self._check_sizes_changed
            )

    def _on_closing(self) -> None:
        self._stop_auto_refresh_sizes()

        if self._auto_save_timer_id is not None:
            try:
                self.root.after_cancel(self._auto_save_timer_id)
            except Exception:
                pass
            self._auto_save_timer_id = None
        self._auto_save_pending = False

        if self.com_manager:
            response = messagebox.askyesnocancel(
                "Đóng ứng dụng",
                "Bạn có muốn lưu thay đổi vào file Excel không?\n\n"
                "Yes: Lưu (Excel vẫn mở)\n"
                "No: Không lưu (Excel vẫn mở)\n"
                "Cancel: Hủy"
            )

            if response is None:
                return

            try:
                self.com_manager.detach(save_changes=response)
                logger.info(f"Đã detach COM manager (save={response}, Excel vẫn chạy)")
            except Exception as e:
                logger.error(f"Lỗi khi detach COM manager: {e}")

        self._save_window_geometry()
        self.root.destroy()

    def _save_window_geometry(self) -> None:
        try:
            self.root.update_idletasks()
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            x = self.root.winfo_x()
            y = self.root.winfo_y()
            self.dialog_config.save_main_window_geometry(width, height, x, y)
        except Exception as e:
            logger.error(f"Lỗi khi lưu geometry cửa sổ chính: {e}")



