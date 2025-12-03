import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Optional, Callable, Tuple
from pathlib import Path
import logging

from excel_automation.excel_com_manager import ExcelCOMManager
from excel_automation.size_filter_config import SizeFilterConfig
from excel_automation.dialog_config_manager import DialogConfigManager
from excel_automation.size_quantity_display_manager import SizeQuantityDisplayManager
from excel_automation.box_list_export_config import BoxListExportConfig
from excel_automation.box_list_export_manager import BoxListExportManager
from ui.size_quantity_input_dialog import SizeQuantityInputDialog

logger = logging.getLogger(__name__)


class ExcelRealtimeController:
    
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

        self.action_frame = ttk.Frame(main_frame)
        self.action_frame.pack(fill=tk.X, pady=(0, 10))

        buttons_config: List[Tuple[str, Callable]] = [
            ("🔍 Quét Sizes", self._scan_sizes),
            ("👁️ Ẩn dòng ngay", self._hide_rows_realtime),
            ("👁️‍🗨️ Hiện tất cả", self._show_all_rows),
            ("📝 Update PO", self._update_po),
            ("🎨 Update Color", self._update_color_code),
            ("📝 Nhập Số Lượng Size", self._input_size_quantities),
            ("📦 Xuất Danh Sách Thùng", self._export_box_list),
        ]

        for text, command in buttons_config:
            btn = ttk.Button(
                self.action_frame,
                text=text,
                command=command,
                width=20
            )
            self.action_buttons.append(btn)

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

        self.sizes_count_label = ttk.Label(
            button_bar,
            text="Chưa quét sizes",
            foreground="gray"
        )
        self.sizes_count_label.pack(side=tk.RIGHT)

        canvas = tk.Canvas(sizes_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(sizes_frame, orient=tk.VERTICAL, command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
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
            
            num_columns = 6
            for idx, size in enumerate(self.available_sizes):
                row = idx // num_columns
                col = idx % num_columns
                
                var = tk.BooleanVar(value=False)
                self.checkboxes[size] = var
                
                cb = ttk.Checkbutton(
                    self.scrollable_frame,
                    text=size,
                    variable=var
                )
                cb.grid(row=row, column=col, sticky=tk.W, padx=10, pady=5)
            
            self.sizes_count_label.config(
                text=f"Tìm thấy {len(self.available_sizes)} sizes",
                foreground="green"
            )
            
            self.status_label.config(
                text=f"Đã quét {len(self.available_sizes)} sizes - "
                f"Cột {self.config.get_column()} "
                f"[{self.config.get_start_row()}:{self.config.get_end_row()}]"
            )
            
            logger.info(f"Đã quét {len(self.available_sizes)} sizes")
            
        except Exception as e:
            logger.error(f"Lỗi khi quét sizes: {e}")
            messagebox.showerror("Lỗi", f"Không thể quét sizes:\n{str(e)}")
            self.status_label.config(text="Lỗi khi quét sizes")
    
    def _select_all_sizes(self) -> None:
        for var in self.checkboxes.values():
            var.set(True)
    
    def _deselect_all_sizes(self) -> None:
        for var in self.checkboxes.values():
            var.set(False)
    
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

                except Exception as e:
                    logger.error(f"Lỗi khi cập nhật mã màu: {e}")
                    messagebox.showerror("Lỗi", f"Không thể cập nhật mã màu:\n{str(e)}")
                    self.status_label.config(text="Lỗi khi cập nhật mã màu")

            ColorCodeUpdateDialog(self.root, current_color, on_save, self.config)

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

                except Exception as e:
                    logger.error(f"Lỗi khi cập nhật PO: {e}")
                    messagebox.showerror("Lỗi", f"Không thể cập nhật PO:\n{str(e)}")
                    self.status_label.config(text="Lỗi khi cập nhật PO")

            POUpdateDialog(self.root, current_po, on_save, self.config)

        except Exception as e:
            logger.error(f"Lỗi khi mở dialog Update PO: {e}")
            messagebox.showerror("Lỗi", f"Không thể mở dialog Update PO:\n{str(e)}")

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

    def _column_number_to_letter(self, col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result

    def _on_closing(self) -> None:
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

