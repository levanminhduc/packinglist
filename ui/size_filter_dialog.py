import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict

from excel_automation.dialog_config_manager import DialogConfigManager


class SizeFilterDialog:

    def __init__(self, parent: tk.Tk, available_sizes: List[str]):
        self.parent = parent
        self.available_sizes = available_sizes
        self.checkboxes: Dict[str, tk.BooleanVar] = {}
        self.selected_sizes: List[str] = []
        self.dialog_config = DialogConfigManager()

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Lọc Size")

        width, height = self.dialog_config.get_dialog_size('size_filter')
        self.dialog.geometry(f"{width}x{height}")
        self.dialog.resizable(True, True)

        self.dialog.transient(parent)
        self.dialog.grab_set()

        self.dialog.protocol("WM_DELETE_WINDOW", self._on_closing)

        self._create_widgets()
        self._center_window()
    
    def _center_window(self) -> None:
        self.dialog.update_idletasks()
        
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        
        self.dialog.geometry(f'{width}x{height}+{x}+{y}')
    
    def _create_widgets(self) -> None:
        header_frame = ttk.Frame(self.dialog)
        header_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(
            header_frame,
            text="Chọn các size muốn hiển thị:",
            font=('Arial', 12, 'bold')
        ).pack(anchor=tk.W)
        
        ttk.Label(
            header_frame,
            text=f"Tìm thấy {len(self.available_sizes)} size khác nhau",
            font=('Arial', 9),
            foreground='gray'
        ).pack(anchor=tk.W, pady=(5, 0))
        
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        ttk.Button(
            button_frame,
            text="✓ Chọn tất cả",
            command=self._select_all
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            button_frame,
            text="✗ Bỏ chọn tất cả",
            command=self._deselect_all
        ).pack(side=tk.LEFT)
        
        ttk.Label(
            button_frame,
            text="Mặc định: Tất cả unchecked (ẩn)",
            font=('Arial', 9),
            foreground='red'
        ).pack(side=tk.RIGHT)
        
        scroll_frame = ttk.LabelFrame(self.dialog, text="Danh sách Size", padding=10)
        scroll_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        canvas = tk.Canvas(scroll_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        num_columns = 6
        for idx, size in enumerate(self.available_sizes):
            row = idx // num_columns
            col = idx % num_columns
            
            var = tk.BooleanVar(value=False)
            self.checkboxes[size] = var
            
            cb = ttk.Checkbutton(
                scrollable_frame,
                text=size,
                variable=var
            )
            cb.grid(row=row, column=col, sticky=tk.W, padx=10, pady=5)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        action_frame = ttk.Frame(self.dialog)
        action_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        ttk.Button(
            action_frame,
            text="Áp dụng",
            command=self._apply_filter,
            width=15
        ).pack(side=tk.RIGHT, padx=(5, 0))
        
        ttk.Button(
            action_frame,
            text="Hủy",
            command=self._on_closing,
            width=15
        ).pack(side=tk.RIGHT)

    def _select_all(self) -> None:
        for var in self.checkboxes.values():
            var.set(True)
    
    def _deselect_all(self) -> None:
        for var in self.checkboxes.values():
            var.set(False)
    
    def _apply_filter(self) -> None:
        self.selected_sizes = [
            size for size, var in self.checkboxes.items()
            if var.get()
        ]

        if not self.selected_sizes:
            response = messagebox.askyesno(
                "Cảnh báo",
                "Bạn chưa chọn size nào!\n\n"
                "Tất cả dòng sẽ bị ẩn.\n\n"
                "Bạn có chắc muốn tiếp tục?"
            )
            if not response:
                return

        self._save_size_and_close()

    def _on_closing(self) -> None:
        self._save_size_and_close()

    def _save_size_and_close(self) -> None:
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            self.dialog_config.save_dialog_size('size_filter', width, height)
        except Exception:
            pass
        finally:
            self.dialog.destroy()
    
    def get_selected_sizes(self) -> List[str]:
        return self.selected_sizes
    
    def show(self) -> bool:
        self.parent.wait_window(self.dialog)
        return len(self.selected_sizes) > 0 or messagebox.askyesno(
            "Xác nhận",
            "Không có size nào được chọn. Tiếp tục?"
        )

