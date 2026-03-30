import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List
import logging

logger = logging.getLogger(__name__)

BASE_HEIGHT = 160
ROW_HEIGHT = 30
GROUP_HEADER_HEIGHT = 35
MIN_HEIGHT = 200
MAX_HEIGHT_RATIO = 0.7
DIALOG_WIDTH = 450


class DuplicateSizeDialog:

    def __init__(self, parent: tk.Tk, duplicate_sizes: Dict[str, List[int]]):
        self.parent = parent
        self.duplicate_sizes = duplicate_sizes
        self.group_checkboxes: Dict[str, Dict[int, tk.BooleanVar]] = {}
        self.rows_to_delete: List[int] = []

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Phát hiện Size trùng")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_skip)

        self._create_widgets()
        self._calculate_and_set_size()
        self._center_window()

    def _calculate_and_set_size(self) -> None:
        total_rows = sum(len(rows) for rows in self.duplicate_sizes.values())
        total_groups = len(self.duplicate_sizes)

        content_height = (
            BASE_HEIGHT
            + total_groups * GROUP_HEADER_HEIGHT
            + total_rows * ROW_HEIGHT
        )

        screen_height = self.parent.winfo_screenheight()
        max_height = int(screen_height * MAX_HEIGHT_RATIO)

        height = max(MIN_HEIGHT, min(content_height, max_height))
        self.dialog.geometry(f"{DIALOG_WIDTH}x{height}")

    def _center_window(self) -> None:
        self.dialog.update_idletasks()
        w = self.dialog.winfo_width()
        h = self.dialog.winfo_height()
        x = (self.parent.winfo_screenwidth() // 2) - (w // 2)
        y = (self.parent.winfo_screenheight() // 2) - (h // 2)
        self.dialog.geometry(f"+{x}+{y}")

    def _create_widgets(self) -> None:
        header_frame = ttk.Frame(self.dialog)
        header_frame.pack(fill=tk.X, padx=10, pady=10)

        total_sizes = len(self.duplicate_sizes)
        total_rows = sum(len(rows) for rows in self.duplicate_sizes.values())

        ttk.Label(
            header_frame,
            text=f"Phát hiện {total_sizes} size trùng ({total_rows} dòng)",
            font=('Arial', 11, 'bold'),
            foreground='#d35400'
        ).pack(anchor=tk.W)

        ttk.Label(
            header_frame,
            text="Check dòng muốn GIỮ, dòng không check sẽ bị XÓA.",
            font=('Arial', 9),
            foreground='gray'
        ).pack(anchor=tk.W, pady=(5, 0))

        scroll_frame = ttk.Frame(self.dialog)
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

        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        )

        from excel_automation.utils import get_size_sort_key
        sorted_sizes = sorted(self.duplicate_sizes.keys(), key=get_size_sort_key)

        for size in sorted_sizes:
            rows = self.duplicate_sizes[size]

            group_frame = ttk.LabelFrame(
                scrollable_frame,
                text=f"Size {size} ({len(rows)} dòng)",
                padding=5
            )
            group_frame.pack(fill=tk.X, padx=5, pady=(5, 0))

            self.group_checkboxes[size] = {}

            for idx, row in enumerate(sorted(rows)):
                var = tk.BooleanVar(value=(idx == 0))
                self.group_checkboxes[size][row] = var

                cb = ttk.Checkbutton(
                    group_frame,
                    text=f"Dòng {row}",
                    variable=var
                )
                cb.pack(anchor=tk.W, padx=10, pady=2)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        action_frame = ttk.Frame(self.dialog)
        action_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Button(
            action_frame,
            text="Xóa dòng trùng",
            command=self._on_delete,
            width=18
        ).pack(side=tk.RIGHT, padx=(5, 0))

        ttk.Button(
            action_frame,
            text="Bỏ qua",
            command=self._on_skip,
            width=12
        ).pack(side=tk.RIGHT)

    def _on_delete(self) -> None:
        for size, row_vars in self.group_checkboxes.items():
            checked_count = sum(1 for var in row_vars.values() if var.get())
            if checked_count == 0:
                messagebox.showwarning(
                    "Cảnh báo",
                    f"Size {size} phải giữ ít nhất 1 dòng!\n"
                    f"Vui lòng check ít nhất 1 dòng cho size {size}.",
                    parent=self.dialog
                )
                return

        rows_to_delete = []
        for size, row_vars in self.group_checkboxes.items():
            for row, var in row_vars.items():
                if not var.get():
                    rows_to_delete.append(row)

        if not rows_to_delete:
            messagebox.showinfo(
                "Thông báo",
                "Tất cả dòng đều được giữ, không có gì để xóa.",
                parent=self.dialog
            )
            self.dialog.destroy()
            return

        confirm = messagebox.askyesno(
            "Xác nhận xóa",
            f"Bạn có chắc muốn XÓA {len(rows_to_delete)} dòng?\n\n"
            f"Dòng sẽ xóa: {', '.join(str(r) for r in sorted(rows_to_delete))}\n\n"
            f"Hành động này không thể hoàn tác!",
            parent=self.dialog
        )

        if confirm:
            self.rows_to_delete = rows_to_delete
            self.dialog.destroy()

    def _on_skip(self) -> None:
        self.rows_to_delete = []
        self.dialog.destroy()

    def get_rows_to_delete(self) -> List[int]:
        return self.rows_to_delete

    def show(self) -> None:
        self.parent.wait_window(self.dialog)
