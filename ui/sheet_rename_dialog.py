import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, List

from excel_automation.dialog_config_manager import DialogConfigManager


class SheetRenameDialog:

    def __init__(self, parent: tk.Tk, default_name: str, existing_names: List[str]):
        self.parent = parent
        self.default_name = default_name
        self.existing_names = [n.lower() for n in existing_names]
        self.result: Optional[str] = None
        self.dialog_config = DialogConfigManager()

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Đổi Tên Sheet")

        width, height = self.dialog_config.get_dialog_size('sheet_rename')
        self.dialog.geometry(f"{width}x{height}")
        self.dialog.resizable(True, True)

        self.dialog.transient(parent)
        self.dialog.grab_set()

        self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)

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
            text="Đổi Tên Sheet Mới",
            font=('Arial', 12, 'bold')
        ).pack(pady=(0, 20))

        name_frame = ttk.Frame(main_frame)
        name_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(name_frame, text="Tên sheet:").pack(side=tk.LEFT)

        self.name_var = tk.StringVar(value=self.default_name)
        self.name_entry = ttk.Entry(name_frame, textvariable=self.name_var, width=30)
        self.name_entry.pack(side=tk.LEFT, padx=(10, 0))
        self.name_entry.select_range(0, tk.END)
        self.name_entry.focus_set()

        self.name_entry.bind('<Return>', lambda e: self._on_save())

        ttk.Label(
            main_frame,
            text="Nhập tên mới hoặc giữ nguyên tên mặc định",
            font=('Arial', 8),
            foreground='gray'
        ).pack(pady=(0, 20))

        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

        ttk.Button(
            button_frame,
            text="Lưu",
            command=self._on_save,
            width=15
        ).pack(side=tk.RIGHT, padx=(5, 0))

        ttk.Button(
            button_frame,
            text="Hủy",
            command=self._on_cancel,
            width=15
        ).pack(side=tk.RIGHT)

    def _on_save(self) -> None:
        name = self.name_var.get().strip()

        if not name:
            messagebox.showerror("Lỗi", "Tên sheet không được rỗng!", parent=self.dialog)
            return

        invalid_chars = ['\\', '/', '?', '*', '[', ']', ':']
        for ch in invalid_chars:
            if ch in name:
                messagebox.showerror(
                    "Lỗi",
                    f"Tên sheet không được chứa ký tự: {' '.join(invalid_chars)}",
                    parent=self.dialog
                )
                return

        if len(name) > 31:
            messagebox.showerror("Lỗi", "Tên sheet không được quá 31 ký tự!", parent=self.dialog)
            return

        if name.lower() in self.existing_names and name != self.default_name:
            messagebox.showerror("Lỗi", f"Sheet '{name}' đã tồn tại!", parent=self.dialog)
            return

        self.result = name
        self._save_and_close()

    def _on_cancel(self) -> None:
        self.result = None
        self._save_and_close()

    def _save_and_close(self) -> None:
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            self.dialog_config.save_dialog_size('sheet_rename', width, height)
        except Exception:
            pass
        finally:
            self.dialog.destroy()

    def show(self) -> Optional[str]:
        self.parent.wait_window(self.dialog)
        return self.result
