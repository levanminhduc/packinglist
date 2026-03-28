import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, Dict, List, Optional
import logging

from excel_automation.pdf_po_parser import PDFPOData
from excel_automation.dialog_config_manager import DialogConfigManager

logger = logging.getLogger(__name__)


class PDFImportDialog:

    def __init__(
        self,
        parent: tk.Tk,
        pdf_data: PDFPOData,
        available_sizes: List[str],
        on_confirm_callback: Callable[[str, str, Dict[str, int]], None]
    ):
        self.parent = parent
        self.pdf_data = pdf_data
        self.available_sizes = available_sizes
        self.on_confirm_callback = on_confirm_callback
        self.dialog_config = DialogConfigManager()

        self.size_checkboxes: Dict[str, tk.BooleanVar] = {}
        self.size_entries: Dict[str, ttk.Entry] = {}
        self.status_labels: Dict[str, ttk.Label] = {}

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("📄 Import PO từ PDF")

        width, height = self.dialog_config.get_dialog_size('pdf_import')
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
        self.dialog.geometry(f"+{x}+{y}")

    def _create_widgets(self) -> None:
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_label = ttk.Label(
            main_frame,
            text=f"File: {self.pdf_data.source_file}",
            foreground="gray"
        )
        file_label.pack(anchor=tk.W, pady=(0, 10))

        info_frame = ttk.LabelFrame(main_frame, text="Thông tin PO", padding=8)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        po_row = ttk.Frame(info_frame)
        po_row.pack(fill=tk.X, pady=2)
        ttk.Label(po_row, text="PO Number:", width=12).pack(side=tk.LEFT)
        self.po_var = tk.StringVar(value=self.pdf_data.po_number)
        ttk.Entry(po_row, textvariable=self.po_var, width=20, font=("Consolas", 11, "bold")).pack(side=tk.LEFT, padx=(5, 0))

        color_row = ttk.Frame(info_frame)
        color_row.pack(fill=tk.X, pady=2)
        ttk.Label(color_row, text="Color Code:", width=12).pack(side=tk.LEFT)
        self.color_var = tk.StringVar(value=self.pdf_data.color_code)
        ttk.Entry(color_row, textvariable=self.color_var, width=20, font=("Consolas", 11, "bold")).pack(side=tk.LEFT, padx=(5, 0))

        total_row = ttk.Frame(info_frame)
        total_row.pack(fill=tk.X, pady=2)
        ttk.Label(total_row, text="Total Qty:", width=12).pack(side=tk.LEFT)
        ttk.Label(total_row, text=f"{self.pdf_data.total_quantity:,}", font=("Consolas", 11, "bold"), foreground="#e65100").pack(side=tk.LEFT, padx=(5, 0))

        self._create_size_table(main_frame)
        self._create_warning_section(main_frame)
        self._create_buttons(main_frame)

    def _create_size_table(self, parent: ttk.Frame) -> None:
        size_frame = ttk.LabelFrame(parent, text="📋 Chi tiết Size — Quantity", padding=8)
        size_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        canvas = tk.Canvas(size_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(size_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable = ttk.Frame(canvas)

        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        header = ttk.Frame(scrollable)
        header.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header, text="☑", width=3, font=("", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="Size", width=10, font=("", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="Qty", width=8, font=("", 9, "bold"), anchor=tk.E).pack(side=tk.LEFT)
        ttk.Label(header, text="Trạng thái", width=16, font=("", 9, "bold")).pack(side=tk.LEFT, padx=(10, 0))

        for size, qty in self.pdf_data.size_quantities.items():
            is_match = size in self.available_sizes
            row = ttk.Frame(scrollable)
            row.pack(fill=tk.X, pady=1)

            var = tk.BooleanVar(value=is_match)
            self.size_checkboxes[size] = var
            ttk.Checkbutton(row, variable=var, command=lambda s=size: self._on_check_changed(s)).pack(side=tk.LEFT)

            size_entry = ttk.Entry(row, width=10, font=("Consolas", 10))
            size_entry.insert(0, size)
            if is_match:
                size_entry.configure(state="readonly")
            else:
                size_entry.bind('<KeyRelease>', lambda e, s=size: self._on_size_edited(s))
            size_entry.pack(side=tk.LEFT, padx=(2, 0))
            self.size_entries[size] = size_entry

            ttk.Label(row, text=str(qty), width=8, anchor=tk.E, font=("Consolas", 10)).pack(side=tk.LEFT)

            status_text = "✅ Khớp" if is_match else "⚠️ Chỉ có trong PDF"
            status_fg = "#2e7d32" if is_match else "#c62828"
            status_label = ttk.Label(row, text=status_text, foreground=status_fg, width=18)
            status_label.pack(side=tk.LEFT, padx=(10, 0))
            self.status_labels[size] = status_label

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_warning_section(self, parent: ttk.Frame) -> None:
        pdf_only = [s for s in self.pdf_data.size_quantities if s not in self.available_sizes]
        excel_only = [s for s in self.available_sizes if s not in self.pdf_data.size_quantities]

        if not pdf_only and not excel_only:
            return

        warn_frame = ttk.Frame(parent)
        warn_frame.pack(fill=tk.X, pady=(0, 10))

        if pdf_only:
            ttk.Label(
                warn_frame,
                text=f"⚠️ {len(pdf_only)} size chỉ có trong PDF: {', '.join(pdf_only)}",
                foreground="#c62828"
            ).pack(anchor=tk.W)

        if excel_only:
            ttk.Label(
                warn_frame,
                text=f"ℹ️ {len(excel_only)} size chỉ có trong Excel: {', '.join(excel_only)}",
                foreground="#1565c0"
            ).pack(anchor=tk.W)

    def _create_buttons(self, parent: ttk.Frame) -> None:
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Hủy", command=self._on_closing, width=15).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="✅ Xác nhận & Ghi", command=self._on_confirm, width=20).pack(side=tk.RIGHT, padx=(0, 5))

    def _on_size_edited(self, original_size: str) -> None:
        entry = self.size_entries[original_size]
        new_value = entry.get().strip()
        label = self.status_labels[original_size]

        if new_value in self.available_sizes:
            label.configure(text="✅ Khớp", foreground="#2e7d32")
            self.size_checkboxes[original_size].set(True)
        else:
            label.configure(text="⚠️ Chỉ có trong PDF", foreground="#c62828")

    def _on_check_changed(self, size: str) -> None:
        pass

    def _on_confirm(self) -> None:
        po = self.po_var.get().strip()
        color = self.color_var.get().strip()

        if not po:
            messagebox.showerror("Lỗi", "PO Number không được để trống", parent=self.dialog)
            return
        if not color:
            messagebox.showerror("Lỗi", "Color Code không được để trống", parent=self.dialog)
            return

        confirmed_sizes: Dict[str, int] = {}
        for original_size, var in self.size_checkboxes.items():
            if var.get():
                entry = self.size_entries[original_size]
                actual_size = entry.get().strip()
                qty = self.pdf_data.size_quantities[original_size]
                confirmed_sizes[actual_size] = qty

        if not confirmed_sizes:
            messagebox.showwarning("Cảnh báo", "Chưa chọn size nào để import", parent=self.dialog)
            return

        total = sum(confirmed_sizes.values())
        confirm_msg = (
            f"Xác nhận import vào Excel:\n\n"
            f"PO: {po}\n"
            f"Color: {color}\n"
            f"Sizes: {len(confirmed_sizes)} size\n"
            f"Total Qty: {total:,}\n\n"
            f"Tiếp tục?"
        )
        if messagebox.askyesno("Xác nhận Import", confirm_msg, parent=self.dialog):
            self._save_size_and_close()
            self.on_confirm_callback(po, color, confirmed_sizes)

    def _on_closing(self) -> None:
        self._save_size_and_close()

    def _save_size_and_close(self) -> None:
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            self.dialog_config.save_dialog_size('pdf_import', width, height)
        except Exception as e:
            logger.error(f"Lỗi khi lưu kích thước dialog: {e}")
        self.dialog.destroy()
