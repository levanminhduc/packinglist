import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable

from excel_automation.color_code_update_manager import ColorCodeUpdateManager
from excel_automation.dialog_config_manager import DialogConfigManager


class ColorCodeUpdateDialog:
    
    def __init__(self, parent: tk.Tk, current_color_code: str, on_save_callback: Callable[[str], None], config):
        self.parent = parent
        self.current_color_code = current_color_code
        self.on_save_callback = on_save_callback
        self.config = config
        self.dialog_config = DialogConfigManager()

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Cập Nhật Mã Màu")

        width, height = self.dialog_config.get_dialog_size('color_code_update')
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
            text="Cập Nhật Mã Màu",
            font=('Arial', 12, 'bold')
        ).pack(pady=(0, 20))
        
        current_frame = ttk.Frame(main_frame)
        current_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(current_frame, text="Mã màu hiện tại:").pack(side=tk.LEFT)
        
        current_color_display = self.current_color_code if self.current_color_code else "(Chưa có)"
        ttk.Label(
            current_frame,
            text=current_color_display,
            font=('Arial', 10, 'bold'),
            foreground='blue'
        ).pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 15))
        
        new_color_frame = ttk.Frame(main_frame)
        new_color_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(new_color_frame, text="Mã màu mới:").pack(side=tk.LEFT)
        
        self.color_var = tk.StringVar(value=self.current_color_code)
        ttk.Entry(new_color_frame, textvariable=self.color_var, width=25).pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Label(
            main_frame,
            text="Nhập mã màu (không được để trống)",
            font=('Arial', 8),
            foreground='gray'
        ).pack(pady=(0, 5))
        
        info_frame = ttk.LabelFrame(main_frame, text="Lưu ý", padding=10)
        info_frame.pack(fill=tk.X, pady=(10, 15))

        ttk.Label(
            info_frame,
            text="• Hệ thống tự động thêm dấu ' phía trước\n"
                 "• Ví dụ: Nhập '0404' → Excel lưu là '0404 (text)\n"
                 "• Đảm bảo giữ nguyên số 0 đầu tiên",
            font=('Arial', 9),
            foreground='#666'
        ).pack(anchor=tk.W)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(15, 0))
        
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
        new_color = self.color_var.get().strip()
        
        is_valid, error_msg = ColorCodeUpdateManager.validate_color_code(None, new_color)
        
        if not is_valid:
            messagebox.showerror("Lỗi", error_msg)
            return
        
        start_row = self.config.get_start_row()
        end_row = self.config.get_end_row()
        
        confirm_msg = (
            f"Bạn có chắc muốn cập nhật mã màu thành:\n\n"
            f"'{new_color}\n\n"
            f"Tất cả dòng từ {start_row} đến {end_row} trong cột E sẽ được cập nhật?"
        )
        
        if messagebox.askyesno("Xác Nhận", confirm_msg):
            self.on_save_callback(new_color)
            self._save_size_and_close()

    def _on_closing(self) -> None:
        self._save_size_and_close()

    def _save_size_and_close(self) -> None:
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            self.dialog_config.save_dialog_size('color_code_update', width, height)
        except Exception:
            pass
        finally:
            self.dialog.destroy()

