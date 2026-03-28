"""
PDF Reader Dialog — Popup dialog để đọc và hiển thị text từ file PDF.

Chạy OCR trên thread riêng để UI không bị đơ.
Dùng root.after() để cập nhật progress — pattern có sẵn trong project.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import logging
from typing import Optional

from excel_automation.pdf_reader import extract_text_from_pdf
from excel_automation.dialog_config_manager import DialogConfigManager

logger = logging.getLogger(__name__)


class PdfReaderDialog:
    """Dialog Tkinter để chọn file PDF và hiển thị text extract được."""

    DIALOG_NAME = "pdf_reader"

    def __init__(self, parent: tk.Tk):
        self.parent = parent
        self.dialog_config = DialogConfigManager()
        self._processing = False
        self._worker_thread: Optional[threading.Thread] = None
        self._destroyed = False

        # Tạo dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Đọc PDF")

        # Kích thước từ config hoặc default 700x500
        width, height = self.dialog_config.get_dialog_size(self.DIALOG_NAME)
        if width < 400:
            width = 700
        if height < 300:
            height = 500
        self.dialog.geometry(f"{width}x{height}")
        self.dialog.resizable(True, True)

        # Transient nhưng không grab_set (cho phép tương tác cửa sổ chính)
        self.dialog.transient(parent)

        self.dialog.protocol("WM_DELETE_WINDOW", self._on_closing)

        self._create_widgets()
        self._center_window()

    def _center_window(self) -> None:
        """Đặt dialog giữa màn hình."""
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")

    def _create_widgets(self) -> None:
        """Tạo tất cả widgets cho dialog."""
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- File chooser row ---
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        self.choose_btn = ttk.Button(
            file_frame,
            text="Chọn file PDF",
            command=self._choose_and_read_pdf,
            width=15
        )
        self.choose_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.file_label = ttk.Label(
            file_frame,
            text="Chưa chọn file",
            foreground="gray"
        )
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- Text result area ---
        self.text_area = ScrolledText(
            main_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            state=tk.DISABLED
        )
        self.text_area.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # --- Status label ---
        self.status_label = ttk.Label(
            main_frame,
            text="Sẵn sàng",
            foreground="gray"
        )
        self.status_label.pack(fill=tk.X, pady=(0, 10))

        # --- Button row ---
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        self.copy_btn = ttk.Button(
            button_frame,
            text="Copy text",
            command=self._copy_text,
            state=tk.DISABLED
        )
        self.copy_btn.pack(side=tk.LEFT)

        ttk.Button(
            button_frame,
            text="Đóng",
            command=self._on_closing
        ).pack(side=tk.RIGHT)

    def _choose_and_read_pdf(self) -> None:
        """Mở file dialog chọn PDF rồi bắt đầu extract text."""
        file_path = filedialog.askopenfilename(
            title="Chọn file PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            parent=self.dialog
        )

        if not file_path:
            return

        self.file_label.config(text=file_path, foreground="black")
        self._start_extraction(file_path)

    def _start_extraction(self, file_path: str) -> None:
        """Bắt đầu extract text trên thread riêng."""
        if self._processing:
            return

        self._processing = True
        self.choose_btn.config(state=tk.DISABLED)
        self.copy_btn.config(state=tk.DISABLED)

        # Xóa text cũ
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete("1.0", tk.END)
        self.text_area.config(state=tk.DISABLED)

        self.status_label.config(text="Đang đọc PDF...", foreground="blue")

        # Chạy extract trên thread riêng
        self._worker_thread = threading.Thread(
            target=self._extract_worker,
            args=(file_path,),
            daemon=True
        )
        self._worker_thread.start()

    def _extract_worker(self, file_path: str) -> None:
        """Worker thread — chạy extract_text_from_pdf."""
        try:
            def on_progress(page_num, total_pages, is_ocr):
                status = f"Đang đọc trang {page_num}/{total_pages}"
                if is_ocr:
                    status += " (OCR)..."
                else:
                    status += "..."
                if not self._destroyed:
                    self.dialog.after(0, self._update_status, status)

            result_text = extract_text_from_pdf(file_path, on_progress=on_progress)
            if not self._destroyed:
                self.dialog.after(0, self._on_extraction_done, result_text)

        except FileNotFoundError as e:
            if not self._destroyed:
                self.dialog.after(0, self._on_extraction_error, f"File không tồn tại: {e}")
        except ValueError as e:
            if not self._destroyed:
                self.dialog.after(0, self._on_extraction_error, f"File không hợp lệ: {e}")
        except Exception as e:
            logger.error(f"Lỗi extract PDF: {e}")
            if not self._destroyed:
                self.dialog.after(0, self._on_extraction_error, f"Lỗi: {e}")

    def _update_status(self, text: str) -> None:
        """Cập nhật status label (gọi từ UI thread)."""
        self.status_label.config(text=text, foreground="blue")

    def _on_extraction_done(self, result_text: str) -> None:
        """Xử lý kết quả extract thành công (gọi từ UI thread)."""
        self._processing = False
        self.choose_btn.config(state=tk.NORMAL)

        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert("1.0", result_text)
        self.text_area.config(state=tk.DISABLED)

        page_count = result_text.count("--- Trang ")
        self.status_label.config(
            text=f"Đã đọc xong — {page_count} trang",
            foreground="green"
        )

        if result_text.strip():
            self.copy_btn.config(state=tk.NORMAL)

    def _on_extraction_error(self, error_msg: str) -> None:
        """Xử lý lỗi extract (gọi từ UI thread)."""
        self._processing = False
        self.choose_btn.config(state=tk.NORMAL)
        self.status_label.config(text=error_msg, foreground="red")
        messagebox.showerror("Lỗi đọc PDF", error_msg, parent=self.dialog)

    def _copy_text(self) -> None:
        """Copy toàn bộ text vào clipboard."""
        text = self.text_area.get("1.0", tk.END).strip()
        if text:
            self.dialog.clipboard_clear()
            self.dialog.clipboard_append(text)
            self.status_label.config(text="Đã copy text vào clipboard!", foreground="green")

    def _on_closing(self) -> None:
        """Xử lý đóng dialog — lưu kích thước, destroy."""
        if not self.dialog.winfo_exists():
            return
        self._destroyed = True
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            self.dialog_config.save_dialog_size(self.DIALOG_NAME, width, height)
        except Exception as e:
            logger.error(f"Lỗi lưu kích thước dialog: {e}")

        self.dialog.destroy()
