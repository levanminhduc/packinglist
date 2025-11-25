import tkinter as tk
from tkinter import ttk, messagebox
from excel_automation.size_filter_config import SizeFilterConfig
from excel_automation.dialog_config_manager import DialogConfigManager


class SizeFilterConfigDialog:
    
    def __init__(self, parent: tk.Tk, config: SizeFilterConfig, max_row: int = None):
        self.parent = parent
        self.config = config
        self.max_row = max_row
        self.dialog_config = DialogConfigManager()

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Cấu Hình Lọc Size")

        width, height = self.dialog_config.get_dialog_size('size_filter_config')
        self.dialog.geometry(f"{width}x{height}")
        self.dialog.resizable(True, True)

        self.dialog.transient(parent)
        self.dialog.grab_set()

        self.dialog.protocol("WM_DELETE_WINDOW", self._on_closing)

        self._create_widgets()
        self._load_current_values()
        self._center_window()
    
    def _center_window(self) -> None:
        self.dialog.update_idletasks()
        
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        
        self.dialog.geometry(f'{width}x{height}+{x}+{y}')
    
    def _create_widgets(self) -> None:
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            main_frame,
            text="Cấu Hình Phạm Vi Quét Size",
            font=('Arial', 12, 'bold')
        ).grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky=tk.W)
        
        ttk.Label(main_frame, text="Tên Sheet:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.sheet_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.sheet_var, width=30).grid(
            row=1, column=1, sticky=tk.W, pady=5, padx=(10, 0)
        )
        
        ttk.Label(main_frame, text="Cột chứa Size:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.column_var = tk.StringVar()
        column_entry = ttk.Entry(main_frame, textvariable=self.column_var, width=10)
        column_entry.grid(row=2, column=1, sticky=tk.W, pady=5, padx=(10, 0))
        ttk.Label(
            main_frame,
            text="(A-ZZZ)",
            font=('Arial', 8),
            foreground='gray'
        ).grid(row=2, column=1, sticky=tk.W, padx=(80, 0))
        
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).grid(
            row=3, column=0, columnspan=2, sticky=tk.EW, pady=15
        )
        
        ttk.Label(
            main_frame,
            text="Phạm Vi Dòng",
            font=('Arial', 10, 'bold')
        ).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(main_frame, text="Dòng bắt đầu:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.start_row_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.start_row_var, width=10).grid(
            row=5, column=1, sticky=tk.W, pady=5, padx=(10, 0)
        )
        
        ttk.Label(main_frame, text="Dòng kết thúc:").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.end_row_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.end_row_var, width=10).grid(
            row=6, column=1, sticky=tk.W, pady=5, padx=(10, 0)
        )
        
        if self.max_row:
            ttk.Label(
                main_frame,
                text=f"(Sheet có {self.max_row} dòng)",
                font=('Arial', 8),
                foreground='blue'
            ).grid(row=6, column=1, sticky=tk.W, padx=(80, 0))
        
        info_frame = ttk.LabelFrame(main_frame, text="Lưu ý", padding=10)
        info_frame.grid(row=7, column=0, columnspan=2, sticky=tk.EW, pady=(15, 0))
        
        ttk.Label(
            info_frame,
            text="• Dòng bắt đầu phải >= 1\n"
                 "• Dòng bắt đầu phải < dòng kết thúc\n"
                 "• Chỉ ẩn/hiện dòng trong phạm vi này\n"
                 "• Dòng ngoài phạm vi luôn hiển thị",
            font=('Arial', 9),
            foreground='#666'
        ).pack(anchor=tk.W)
        
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        ttk.Button(
            button_frame,
            text="Lưu",
            command=self._save_config,
            width=15
        ).pack(side=tk.RIGHT, padx=(5, 0))
        
        ttk.Button(
            button_frame,
            text="Hủy",
            command=self._on_closing,
            width=15
        ).pack(side=tk.RIGHT)
        
        ttk.Button(
            button_frame,
            text="Reset Mặc Định",
            command=self._reset_to_defaults,
            width=15
        ).pack(side=tk.LEFT)
    
    def _load_current_values(self) -> None:
        self.sheet_var.set(self.config.get_sheet_name())
        self.column_var.set(self.config.get_column())
        self.start_row_var.set(str(self.config.get_start_row()))
        self.end_row_var.set(str(self.config.get_end_row()))
    
    def _save_config(self) -> None:
        try:
            sheet_name = self.sheet_var.get().strip()
            column = self.column_var.get().strip().upper()
            start_row = int(self.start_row_var.get())
            end_row = int(self.end_row_var.get())
            
            if not sheet_name:
                messagebox.showerror("Lỗi", "Tên sheet không được rỗng!")
                return
            
            if not column or len(column) > 3:
                messagebox.showerror("Lỗi", "Cột phải là 1-3 ký tự (A-ZZZ)!")
                return
            
            if start_row < 1:
                messagebox.showerror("Lỗi", "Dòng bắt đầu phải >= 1!")
                return
            
            if start_row >= end_row:
                messagebox.showerror("Lỗi", "Dòng bắt đầu phải < dòng kết thúc!")
                return
            
            if self.max_row and end_row > self.max_row:
                messagebox.showerror(
                    "Lỗi",
                    f"Dòng kết thúc ({end_row}) vượt quá số dòng thực tế ({self.max_row})!"
                )
                return
            
            self.config.update_config(column, start_row, end_row, sheet_name)
            
            messagebox.showinfo(
                "Thành Công",
                f"Đã lưu cấu hình:\n\n"
                f"Sheet: {sheet_name}\n"
                f"Cột: {column}\n"
                f"Phạm vi: {start_row} - {end_row}"
            )
            self._save_size_and_close()
            
        except ValueError as e:
            messagebox.showerror("Lỗi", f"Giá trị không hợp lệ:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi lưu cấu hình:\n{str(e)}")
    
    def _reset_to_defaults(self) -> None:
        if messagebox.askyesno(
            "Xác Nhận",
            "Bạn có chắc muốn reset về cấu hình mặc định?\n\n"
            "Cột: F\n"
            "Phạm vi: 19 - 59\n"
            "Sheet: Sheet1"
        ):
            self.config.reset_to_defaults()
            self._load_current_values()
            messagebox.showinfo("Thành Công", "Đã reset về cấu hình mặc định")

    def _on_closing(self) -> None:
        self._save_size_and_close()

    def _save_size_and_close(self) -> None:
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            self.dialog_config.save_dialog_size('size_filter_config', width, height)
        except Exception:
            pass
        finally:
            self.dialog.destroy()

