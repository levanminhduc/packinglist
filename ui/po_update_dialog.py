import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable

from excel_automation.po_update_manager import POUpdateManager
from excel_automation.dialog_config_manager import DialogConfigManager


class POUpdateDialog:
    
    def __init__(self, parent: tk.Tk, current_po: str, on_save_callback: Callable[[str], None], config):
        self.parent = parent
        self.current_po = current_po
        self.on_save_callback = on_save_callback
        self.config = config
        self.dialog_config = DialogConfigManager()

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Cập Nhật PO")

        width, height = self.dialog_config.get_dialog_size('po_update')
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
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            main_frame,
            text="Cập Nhật Purchase Order",
            font=('Arial', 12, 'bold')
        ).pack(pady=(0, 20))
        
        current_frame = ttk.Frame(main_frame)
        current_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(current_frame, text="PO hiện tại:").pack(side=tk.LEFT)
        
        current_po_display = self.current_po if self.current_po else "(Chưa có)"
        ttk.Label(
            current_frame,
            text=current_po_display,
            font=('Arial', 10, 'bold'),
            foreground='blue'
        ).pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 15))
        
        new_po_frame = ttk.Frame(main_frame)
        new_po_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(new_po_frame, text="PO mới:").pack(side=tk.LEFT)
        
        self.po_var = tk.StringVar(value=self.current_po)
        ttk.Entry(new_po_frame, textvariable=self.po_var, width=25).pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Label(
            main_frame,
            text="Nhập giá trị mới (không được để trống)",
            font=('Arial', 8),
            foreground='gray'
        ).pack(pady=(0, 20))
        
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        ttk.Button(
            button_frame,
            text="Lưu",
            command=self._save,
            width=15
        ).pack(side=tk.RIGHT, padx=(5, 0))
        
        ttk.Button(
            button_frame,
            text="Hủy",
            command=self._on_closing,
            width=15
        ).pack(side=tk.RIGHT)
    
    def _save(self) -> None:
        new_po = self.po_var.get().strip()
        
        is_valid, error_msg = POUpdateManager.validate_po(None, new_po)
        
        if not is_valid:
            messagebox.showerror("Lỗi", error_msg)
            return
        
        start_row = self.config.get_start_row()
        end_row = self.config.get_end_row()
        
        confirm_msg = (
            f"Bạn có chắc muốn cập nhật PO thành:\n\n"
            f"{new_po}\n\n"
            f"Tất cả dòng từ {start_row} đến {end_row} trong cột A sẽ được cập nhật?"
        )
        
        if messagebox.askyesno("Xác Nhận", confirm_msg):
            self.on_save_callback(new_po)
            self._save_size_and_close()

    def _on_closing(self) -> None:
        self._save_size_and_close()

    def _save_size_and_close(self) -> None:
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            self.dialog_config.save_dialog_size('po_update', width, height)
        except Exception:
            pass
        finally:
            self.dialog.destroy()

