import tkinter as tk
from tkinter import ttk
from typing import List, Optional, Callable
import logging

logger = logging.getLogger(__name__)


class BoxListExportProgressDialog:

    STEPS = [
        "Đọc dữ liệu thùng từ Excel",
        "Phân tích & gộp sizes",
        "Tạo sheet mới",
        "Ghi danh sách thùng vào sheet",
        "Copy vào clipboard",
        "Hoàn tất",
    ]

    STEP_WEIGHTS = [30, 15, 15, 25, 10, 5]

    def __init__(self, parent: tk.Tk):
        self.parent = parent
        self.current_step = 0
        self.step_labels: List[ttk.Label] = []
        self.retry_callback: Optional[Callable] = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Đang xuất danh sách thùng...")
        self.dialog.geometry("420x350")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)

        self._create_widgets()
        self._center_window()

    def _center_window(self) -> None:
        self.dialog.update_idletasks()
        w = self.dialog.winfo_width()
        h = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (w // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (h // 2)
        self.dialog.geometry(f"+{x}+{y}")

    def _create_widgets(self) -> None:
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(
            main_frame, variable=self.progress_var,
            maximum=100, length=350, mode='determinate'
        )
        self.progress_bar.pack(pady=(0, 5))

        self.percent_label = ttk.Label(main_frame, text="0%", font=("", 11, "bold"))
        self.percent_label.pack(pady=(0, 15))

        steps_frame = ttk.Frame(main_frame)
        steps_frame.pack(fill=tk.BOTH, expand=True)

        for step_text in self.STEPS:
            label = ttk.Label(steps_frame, text=f"  \u2b1a  {step_text}", foreground="gray")
            label.pack(anchor=tk.W, pady=2)
            self.step_labels.append(label)

        self.btn_frame = ttk.Frame(main_frame)
        self.btn_frame.pack(fill=tk.X, pady=(15, 0))

        self.error_label = ttk.Label(main_frame, text="", foreground="#c62828", wraplength=360)

    def start_step(self, step_index: int) -> None:
        self.current_step = step_index
        percent = sum(self.STEP_WEIGHTS[:step_index])
        self.progress_var.set(percent)
        self.percent_label.configure(text=f"{percent}%")

        for i, label in enumerate(self.step_labels):
            if i < step_index:
                label.configure(text=f"  \u2705  {self.STEPS[i]}", foreground="#2e7d32")
            elif i == step_index:
                label.configure(text=f"  \U0001f504  Đang {self.STEPS[i].lower()}...", foreground="#1565c0")
            else:
                label.configure(text=f"  \u2b1a  {self.STEPS[i]}", foreground="gray")

        self.dialog.update()

    def complete_step(self, step_index: int) -> None:
        self.step_labels[step_index].configure(
            text=f"  \u2705  {self.STEPS[step_index]}", foreground="#2e7d32"
        )
        percent = sum(self.STEP_WEIGHTS[:step_index + 1])
        self.progress_var.set(percent)
        self.percent_label.configure(text=f"{percent}%")
        self.dialog.update()

    def finish(self) -> None:
        self.progress_var.set(100)
        self.percent_label.configure(text="100%")
        for i, label in enumerate(self.step_labels):
            label.configure(text=f"  \u2705  {self.STEPS[i]}", foreground="#2e7d32")
        self.dialog.update()
        self.parent.after(1000, self.dialog.destroy)

    def show_error(self, step_index: int, error_msg: str, retry_callback: Callable) -> None:
        self.step_labels[step_index].configure(
            text=f"  \u274c  {self.STEPS[step_index]}", foreground="#c62828"
        )
        self.error_label.configure(text=f"Lỗi: {error_msg}")
        self.error_label.pack(pady=(10, 0))

        self.retry_callback = retry_callback

        for widget in self.btn_frame.winfo_children():
            widget.destroy()

        ttk.Button(
            self.btn_frame, text="\U0001f504 Thử lại",
            command=self._retry, width=15
        ).pack(side=tk.LEFT)
        ttk.Button(
            self.btn_frame, text="Đóng",
            command=self.dialog.destroy, width=15
        ).pack(side=tk.RIGHT)

        self.dialog.protocol("WM_DELETE_WINDOW", self.dialog.destroy)
        self.dialog.update()

    def close(self) -> None:
        self.dialog.destroy()

    def _retry(self) -> None:
        self.error_label.pack_forget()
        for widget in self.btn_frame.winfo_children():
            widget.destroy()
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        if self.retry_callback:
            self.retry_callback()
