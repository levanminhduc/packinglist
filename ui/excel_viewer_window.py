import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import Optional
import pandas as pd
import logging

from excel_automation import ExcelReader, DataValidator, ValidationResult
from excel_automation.size_filter import SizeFilterManager
from excel_automation.size_filter_config import SizeFilterConfig
from ui.ui_config import UIConfig
from ui.size_filter_dialog import SizeFilterDialog
from ui.size_filter_config_dialog import SizeFilterConfigDialog

logger = logging.getLogger(__name__)


class ExcelViewerWindow:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.config = UIConfig()
        self.size_filter_config = SizeFilterConfig()
        self.current_file: Optional[str] = None
        self.df: Optional[pd.DataFrame] = None
        self.sheet_names: list = []
        self.current_sheet: str = None
        self.all_sheets_data: dict = {}
        self.validation_result: Optional[ValidationResult] = None
        self.validator: Optional[DataValidator] = None

        self._setup_window()
        self._create_menu()
        self._create_toolbar()
        self._create_sheet_tabs()
        self._create_table()
        self._create_statusbar()

        self._load_last_file()
    
    def _setup_window(self) -> None:
        self.root.title("Excel Viewer - Đọc File Excel")
        
        geometry = self.config.get_window_geometry()
        self.root.geometry(geometry)
        
        if self.config.get('window.maximized', False):
            self.root.state('zoomed')
        
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _create_menu(self) -> None:
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Mở File Excel...", command=self._open_file, accelerator="Ctrl+O")
        file_menu.add_separator()
        
        recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="File Gần Đây", menu=recent_menu)
        self._update_recent_menu(recent_menu)
        
        file_menu.add_separator()
        file_menu.add_command(label="Thoát", command=self._on_closing, accelerator="Ctrl+Q")
        
        validation_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Validation", menu=validation_menu)
        validation_menu.add_command(label="Load Rules từ JSON...", command=self._load_validation_rules, accelerator="Ctrl+L")
        validation_menu.add_command(label="Validate Dữ Liệu", command=self._validate_data, accelerator="Ctrl+V")
        validation_menu.add_separator()
        validation_menu.add_command(label="Xem Kết Quả Validation", command=self._show_validation_results)
        validation_menu.add_command(label="Export Báo Cáo Lỗi...", command=self._export_error_report)
        validation_menu.add_command(label="Xóa Validation", command=self._clear_validation)

        size_filter_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Lọc Size", menu=size_filter_menu)
        size_filter_menu.add_command(label="Lọc theo Size...", command=self._open_size_filter, accelerator="Ctrl+F")
        size_filter_menu.add_command(label="Cấu hình Lọc Size...", command=self._open_size_filter_config)
        size_filter_menu.add_separator()
        size_filter_menu.add_command(label="Real-Time Controller...", command=self._open_realtime_controller, accelerator="Ctrl+R")
        size_filter_menu.add_separator()
        size_filter_menu.add_command(label="Reset Lọc Size", command=self._reset_size_filter)

        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Cài Đặt", menu=settings_menu)
        settings_menu.add_command(label="Tùy Chỉnh Giao Diện...", command=self._open_settings)
        settings_menu.add_command(label="Reset Về Mặc Định", command=self._reset_settings)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Trợ Giúp", menu=help_menu)
        help_menu.add_command(label="Về Chương Trình", command=self._show_about)
        
        self.root.bind('<Control-o>', lambda e: self._open_file())
        self.root.bind('<Control-q>', lambda e: self._on_closing())
        self.root.bind('<Control-l>', lambda e: self._load_validation_rules())
        self.root.bind('<Control-v>', lambda e: self._validate_data())
        self.root.bind('<Control-f>', lambda e: self._open_size_filter())
        self.root.bind('<Control-r>', lambda e: self._open_realtime_controller())
    
    def _create_toolbar(self) -> None:
        toolbar = ttk.Frame(self.root)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        
        ttk.Button(
            toolbar,
            text="📂 Mở File",
            command=self._open_file
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            toolbar,
            text="🔄 Tải Lại",
            command=self._reload_file
        ).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        ttk.Button(
            toolbar,
            text="📋 Load Rules",
            command=self._load_validation_rules
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            toolbar,
            text="✓ Validate",
            command=self._validate_data
        ).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        ttk.Button(
            toolbar,
            text="🔍 Lọc Size",
            command=self._open_size_filter
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            toolbar,
            text="⚙️ Config Size",
            command=self._open_size_filter_config
        ).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        ttk.Button(
            toolbar,
            text="⚙️ Cài Đặt",
            command=self._open_settings
        ).pack(side=tk.LEFT, padx=2)

        self.file_label = ttk.Label(toolbar, text="Chưa mở file nào", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=10)

        self.validation_label = ttk.Label(toolbar, text="", foreground="gray")
        self.validation_label.pack(side=tk.RIGHT, padx=10)

    def _create_sheet_tabs(self) -> None:
        self.sheet_tab_frame = ttk.Frame(self.root)
        self.sheet_tab_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(0, 5))

        self.sheet_notebook = ttk.Notebook(self.sheet_tab_frame)
        self.sheet_notebook.pack(fill=tk.BOTH, expand=True)
        self.sheet_notebook.bind('<<NotebookTabChanged>>', self._on_sheet_changed)

    def _create_table(self) -> None:
        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.tree = ttk.Treeview(
            table_frame,
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
            show='tree headings'
        )
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        
        self._apply_table_config()
    
    def _create_statusbar(self) -> None:
        statusbar = ttk.Frame(self.root)
        statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_label = ttk.Label(statusbar, text="Sẵn sàng", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.row_count_label = ttk.Label(statusbar, text="0 dòng", relief=tk.SUNKEN)
        self.row_count_label.pack(side=tk.RIGHT, padx=5)
    
    def _apply_table_config(self) -> None:
        table_config = self.config.get_table_config()
        
        font_family = table_config.get('font_family', 'Arial')
        font_size = table_config.get('font_size', 10)
        
        style = ttk.Style()
        style.configure('Treeview', font=(font_family, font_size))
        style.configure('Treeview.Heading', font=(font_family, font_size, 'bold'))
    
    def _open_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Chọn File Excel",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("All Files", "*.*")
            ]
        )
        
        if file_path:
            self._load_file(file_path)
    
    def _load_file(self, file_path: str) -> None:
        try:
            self.status_label.config(text=f"Đang đọc file: {Path(file_path).name}...")
            self.root.update()

            reader = ExcelReader(file_path)
            self.all_sheets_data = reader.read_all_sheets()
            self.sheet_names = list(self.all_sheets_data.keys())

            self.current_file = file_path
            self.config.add_recent_file(file_path)

            self._update_sheet_tabs()

            if self.sheet_names:
                self.current_sheet = self.sheet_names[0]
                self.df = self.all_sheets_data[self.current_sheet]
                self._display_dataframe()

            self.file_label.config(text=f"📄 {Path(file_path).name}", foreground="black")
            self.status_label.config(text=f"Đã tải: {Path(file_path).name} - {len(self.sheet_names)} sheets")

            logger.info(f"Đã mở file: {file_path} với {len(self.sheet_names)} sheets")

        except Exception as e:
            logger.error(f"Lỗi khi mở file: {e}")
            messagebox.showerror("Lỗi", f"Không thể mở file:\n{str(e)}")
            self.status_label.config(text="Lỗi khi đọc file")
    
    def _display_dataframe(self) -> None:
        if self.df is None:
            return
        
        self.tree.delete(*self.tree.get_children())
        
        columns = list(self.df.columns)
        self.tree['columns'] = columns
        self.tree.column('#0', width=50, minwidth=50, stretch=tk.NO)
        self.tree.heading('#0', text='#')
        
        column_width = self.config.get('table.column_width', 150)
        
        for col in columns:
            self.tree.column(col, width=column_width, minwidth=50)
            self.tree.heading(col, text=str(col))
        
        for idx, row in self.df.iterrows():
            values = [str(val) if pd.notna(val) else '' for val in row]
            self.tree.insert('', tk.END, text=str(idx + 1), values=values)
        
        self.row_count_label.config(text=f"{len(self.df)} dòng")
    
    def _update_sheet_tabs(self) -> None:
        for tab in self.sheet_notebook.tabs():
            self.sheet_notebook.forget(tab)

        for sheet_name in self.sheet_names:
            frame = ttk.Frame(self.sheet_notebook)
            self.sheet_notebook.add(frame, text=sheet_name)

    def _on_sheet_changed(self, event) -> None:
        try:
            selected_tab = self.sheet_notebook.index(self.sheet_notebook.select())
            if 0 <= selected_tab < len(self.sheet_names):
                self.current_sheet = self.sheet_names[selected_tab]
                self.df = self.all_sheets_data[self.current_sheet]
                self._display_dataframe()
                self.status_label.config(text=f"Sheet: {self.current_sheet} - {len(self.df)} dòng")
                logger.info(f"Chuyển sang sheet: {self.current_sheet}")
        except Exception as e:
            logger.error(f"Lỗi khi chuyển sheet: {e}")

    def _reload_file(self) -> None:
        if self.current_file:
            self._load_file(self.current_file)
        else:
            messagebox.showinfo("Thông Báo", "Chưa mở file nào để tải lại")
    
    def _load_last_file(self) -> None:
        last_file = self.config.get('last_opened_file')
        if last_file and Path(last_file).exists():
            self._load_file(last_file)
    
    def _update_recent_menu(self, menu: tk.Menu) -> None:
        menu.delete(0, tk.END)
        recent_files = self.config.get_recent_files()
        
        if not recent_files:
            menu.add_command(label="(Trống)", state=tk.DISABLED)
        else:
            for file_path in recent_files:
                if Path(file_path).exists():
                    menu.add_command(
                        label=Path(file_path).name,
                        command=lambda f=file_path: self._load_file(f)
                    )
    
    def _open_settings(self) -> None:
        from ui.settings_dialog import SettingsDialog
        dialog = SettingsDialog(self.root, self.config)
        self.root.wait_window(dialog.dialog)
        
        self._apply_table_config()
        if self.df is not None:
            self._display_dataframe()
    
    def _reset_settings(self) -> None:
        if messagebox.askyesno("Xác Nhận", "Bạn có chắc muốn reset tất cả cài đặt về mặc định?"):
            self.config.reset_to_defaults()
            self._apply_table_config()
            if self.df is not None:
                self._display_dataframe()
            messagebox.showinfo("Thành Công", "Đã reset cài đặt về mặc định")
    
    def _show_about(self) -> None:
        messagebox.showinfo(
            "Về Chương Trình",
            "Excel Viewer v1.0\n\n"
            "Ứng dụng đọc và hiển thị file Excel\n"
            "với khả năng tùy chỉnh giao diện\n\n"
            "© 2025 Excel Automation"
        )
    
    def _load_validation_rules(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Chọn File Validation Rules (JSON)",
            filetypes=[
                ("JSON Files", "*.json"),
                ("All Files", "*.*")
            ],
            initialdir="data/validation_rules"
        )

        if file_path:
            try:
                self.validator = DataValidator.from_json(file_path)
                rules_count = sum(len(rules) for rules in self.validator.rules.values())
                self.validation_label.config(
                    text=f"📋 Rules: {len(self.validator.rules)} cột, {rules_count} rules",
                    foreground="blue"
                )
                messagebox.showinfo(
                    "Thành Công",
                    f"Đã load {rules_count} validation rules cho {len(self.validator.rules)} cột"
                )
                logger.info(f"Đã load validation rules từ: {file_path}")
            except Exception as e:
                logger.error(f"Lỗi khi load validation rules: {e}")
                messagebox.showerror("Lỗi", f"Không thể load validation rules:\n{str(e)}")

    def _validate_data(self) -> None:
        if self.df is None:
            messagebox.showwarning("Cảnh Báo", "Chưa mở file nào để validate")
            return

        if self.validator is None:
            response = messagebox.askyesno(
                "Chưa Load Rules",
                "Chưa load validation rules. Bạn có muốn load rules trước không?"
            )
            if response:
                self._load_validation_rules()
                if self.validator is None:
                    return
            else:
                return

        try:
            self.status_label.config(text="Đang validate dữ liệu...")
            self.root.update()

            self.validation_result = self.validator.validate_dataframe(self.df)

            if self.validation_result.is_valid:
                self.validation_label.config(
                    text=f"✅ Valid: {self.validation_result.total_rows} dòng",
                    foreground="green"
                )
                messagebox.showinfo(
                    "Validation Thành Công",
                    f"✅ Tất cả {self.validation_result.total_rows} dòng dữ liệu đều hợp lệ!"
                )
            else:
                self.validation_label.config(
                    text=f"❌ Lỗi: {self.validation_result.error_count}/{self.validation_result.total_rows}",
                    foreground="red"
                )
                self._highlight_validation_errors()

                response = messagebox.askyesno(
                    "Validation Thất Bại",
                    f"❌ Tìm thấy {self.validation_result.error_count} lỗi trong {self.validation_result.total_rows} dòng\n\n"
                    f"Dòng hợp lệ: {self.validation_result.summary['valid_rows']}\n\n"
                    "Bạn có muốn xem chi tiết không?"
                )
                if response:
                    self._show_validation_results()

            self.status_label.config(text=f"Validation hoàn thành: {self.validation_result.error_count} lỗi")
            logger.info(f"Validation hoàn thành: {self.validation_result.error_count} lỗi")

        except Exception as e:
            logger.error(f"Lỗi khi validate: {e}")
            messagebox.showerror("Lỗi", f"Không thể validate dữ liệu:\n{str(e)}")
            self.status_label.config(text="Lỗi khi validate")

    def _highlight_validation_errors(self) -> None:
        if self.validation_result is None or self.validation_result.is_valid:
            return

        error_rows = set()
        for error in self.validation_result.errors:
            error_rows.add(error.row_index - 2)

        for item in self.tree.get_children():
            row_idx = int(self.tree.item(item)['text']) - 1
            if row_idx in error_rows:
                self.tree.item(item, tags=('error',))

        self.tree.tag_configure('error', background='#FFFF99', foreground='#CC0000')

    def _show_validation_results(self) -> None:
        if self.validation_result is None:
            messagebox.showinfo("Thông Báo", "Chưa có kết quả validation")
            return

        result_window = tk.Toplevel(self.root)
        result_window.title("Kết Quả Validation")
        result_window.geometry("800x600")

        frame = ttk.Frame(result_window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        summary_frame = ttk.LabelFrame(frame, text="Tổng Quan", padding=10)
        summary_frame.pack(fill=tk.X, pady=(0, 10))

        summary = self.validation_result.summary
        ttk.Label(summary_frame, text=f"Tổng số dòng: {self.validation_result.total_rows}").pack(anchor=tk.W)
        ttk.Label(summary_frame, text=f"Dòng hợp lệ: {summary['valid_rows']}").pack(anchor=tk.W)
        ttk.Label(summary_frame, text=f"Số lỗi: {self.validation_result.error_count}").pack(anchor=tk.W)

        status_text = "✅ PASS" if self.validation_result.is_valid else "❌ FAIL"
        status_color = "green" if self.validation_result.is_valid else "red"
        ttk.Label(summary_frame, text=f"Trạng thái: {status_text}", foreground=status_color).pack(anchor=tk.W)

        if not self.validation_result.is_valid:
            errors_frame = ttk.LabelFrame(frame, text="Chi Tiết Lỗi", padding=10)
            errors_frame.pack(fill=tk.BOTH, expand=True)

            scrollbar = ttk.Scrollbar(errors_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            error_tree = ttk.Treeview(
                errors_frame,
                columns=('Dòng', 'Cột', 'Giá Trị', 'Quy Tắc', 'Lỗi'),
                show='headings',
                yscrollcommand=scrollbar.set
            )
            error_tree.pack(fill=tk.BOTH, expand=True)
            scrollbar.config(command=error_tree.yview)

            error_tree.heading('Dòng', text='Dòng')
            error_tree.heading('Cột', text='Cột')
            error_tree.heading('Giá Trị', text='Giá Trị')
            error_tree.heading('Quy Tắc', text='Quy Tắc')
            error_tree.heading('Lỗi', text='Thông Báo Lỗi')

            error_tree.column('Dòng', width=60)
            error_tree.column('Cột', width=100)
            error_tree.column('Giá Trị', width=100)
            error_tree.column('Quy Tắc', width=100)
            error_tree.column('Lỗi', width=400)

            for error in self.validation_result.errors:
                error_tree.insert('', tk.END, values=(
                    error.row_index,
                    error.column,
                    str(error.value)[:50],
                    error.rule,
                    error.message
                ))

        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(
            button_frame,
            text="Export Báo Cáo",
            command=self._export_error_report
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="Đóng",
            command=result_window.destroy
        ).pack(side=tk.RIGHT, padx=5)

    def _export_error_report(self) -> None:
        if self.validation_result is None:
            messagebox.showinfo("Thông Báo", "Chưa có kết quả validation")
            return

        if self.validation_result.is_valid:
            messagebox.showinfo("Thông Báo", "Không có lỗi để export")
            return

        file_path = filedialog.asksaveasfilename(
            title="Lưu Báo Cáo Lỗi",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialdir="data/output"
        )

        if file_path:
            try:
                self.validator.generate_error_report(self.validation_result, file_path)
                messagebox.showinfo(
                    "Thành Công",
                    f"Đã export báo cáo lỗi tại:\n{file_path}"
                )
                logger.info(f"Đã export báo cáo lỗi: {file_path}")
            except Exception as e:
                logger.error(f"Lỗi khi export báo cáo: {e}")
                messagebox.showerror("Lỗi", f"Không thể export báo cáo:\n{str(e)}")

    def _clear_validation(self) -> None:
        self.validation_result = None
        self.validator = None
        self.validation_label.config(text="", foreground="gray")

        for item in self.tree.get_children():
            self.tree.item(item, tags=())

        messagebox.showinfo("Thành Công", "Đã xóa validation")
        logger.info("Đã xóa validation")

    def _open_size_filter(self) -> None:
        if not self.current_file:
            messagebox.showwarning("Cảnh Báo", "Vui lòng mở file Excel trước!")
            return

        try:
            with SizeFilterManager(self.current_file, self.size_filter_config) as manager:
                available_sizes = manager.scan_sizes()

                if not available_sizes:
                    messagebox.showinfo(
                        "Thông Báo",
                        f"Không tìm thấy size nào trong cột {self.size_filter_config.get_column()} "
                        f"[{self.size_filter_config.get_start_row()}:{self.size_filter_config.get_end_row()}]"
                    )
                    return

                dialog = SizeFilterDialog(self.root, available_sizes)
                dialog.show()

                selected_sizes = dialog.get_selected_sizes()

                if selected_sizes or messagebox.askyesno("Xác nhận", "Không có size nào được chọn. Tiếp tục?"):
                    hidden_count = manager.apply_size_filter(selected_sizes)
                    manager.save()

                    messagebox.showinfo(
                        "Thành Công",
                        f"Đã áp dụng lọc size:\n\n"
                        f"- Số size được chọn: {len(selected_sizes)}\n"
                        f"- Số dòng bị ẩn: {hidden_count}\n\n"
                        f"Vui lòng tải lại file để xem kết quả."
                    )

                    self._reload_file()
                    logger.info(f"Đã lọc size: {len(selected_sizes)} sizes, ẩn {hidden_count} dòng")

        except Exception as e:
            logger.error(f"Lỗi khi lọc size: {e}")
            messagebox.showerror("Lỗi", f"Không thể lọc size:\n{str(e)}")

    def _open_size_filter_config(self) -> None:
        try:
            max_row = None
            if self.current_file:
                with SizeFilterManager(self.current_file, self.size_filter_config) as manager:
                    manager._load_workbook()
                    max_row = manager.ws.max_row

            dialog = SizeFilterConfigDialog(self.root, self.size_filter_config, max_row)

        except Exception as e:
            logger.error(f"Lỗi khi mở config size filter: {e}")
            messagebox.showerror("Lỗi", f"Không thể mở cấu hình:\n{str(e)}")

    def _reset_size_filter(self) -> None:
        if not self.current_file:
            messagebox.showwarning("Cảnh Báo", "Vui lòng mở file Excel trước!")
            return

        if not messagebox.askyesno(
            "Xác Nhận",
            "Bạn có chắc muốn hiện lại tất cả các dòng đã bị ẩn?"
        ):
            return

        try:
            with SizeFilterManager(self.current_file, self.size_filter_config) as manager:
                manager.reset_all_rows()
                manager.save()

                messagebox.showinfo(
                    "Thành Công",
                    f"Đã hiện lại tất cả dòng từ {self.size_filter_config.get_start_row()} "
                    f"đến {self.size_filter_config.get_end_row()}"
                )

                self._reload_file()
                logger.info("Đã reset size filter")

        except Exception as e:
            logger.error(f"Lỗi khi reset size filter: {e}")
            messagebox.showerror("Lỗi", f"Không thể reset:\n{str(e)}")

    def _open_realtime_controller(self) -> None:
        try:
            import subprocess
            import sys

            controller_script = Path(__file__).parent.parent / "excel_realtime_controller.py"

            if not controller_script.exists():
                messagebox.showerror(
                    "Lỗi",
                    f"Không tìm thấy file excel_realtime_controller.py"
                )
                return

            subprocess.Popen([sys.executable, str(controller_script)])

            messagebox.showinfo(
                "Thông Báo",
                "Đã mở Excel Real-Time Controller trong cửa sổ mới!\n\n"
                "Real-Time Controller cho phép bạn:\n"
                "- Điều khiển Excel trực tiếp qua COM\n"
                "- Ẩn/hiện dòng real-time không cần reload\n"
                "- Chọn sheet động và quét sizes tự động"
            )

            logger.info("Đã mở Real-Time Controller")

        except Exception as e:
            logger.error(f"Lỗi khi mở Real-Time Controller: {e}")
            messagebox.showerror("Lỗi", f"Không thể mở Real-Time Controller:\n{str(e)}")

    def _on_closing(self) -> None:
        try:
            geometry = self.root.geometry()
            self.config.set_window_geometry(geometry)

            is_maximized = self.root.state() == 'zoomed'
            self.config.set('window.maximized', is_maximized)

            logger.info("Đã lưu cấu hình trước khi đóng")
        except Exception as e:
            logger.error(f"Lỗi khi lưu cấu hình: {e}")
        finally:
            self.root.destroy()

