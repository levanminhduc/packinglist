import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict, Optional, TYPE_CHECKING
import logging
import math
import re

from excel_automation.utils import get_size_sort_key
from excel_automation.carton_allocation_calculator import (
    CartonAllocationCalculator,
    AllocationResult
)
from excel_automation.dialog_config_manager import DialogConfigManager

if TYPE_CHECKING:
    from win32com.client import CDispatch

logger = logging.getLogger(__name__)


class SizeQuantityInputDialog:

    def __init__(self, parent: tk.Tk, selected_sizes: List[str], current_quantities: Optional[Dict[str, Optional[int]]] = None, worksheet: Optional['CDispatch'] = None):
        self.parent = parent
        self.selected_sizes = sorted(selected_sizes, key=get_size_sort_key)
        self.current_quantities = current_quantities or {}
        self.quantity_entries: Dict[str, tk.Entry] = {}
        self.quantities: Dict[str, int] = {}
        self.canvas: Optional[tk.Canvas] = None

        self.worksheet = worksheet
        self.total_qty: Optional[int] = None
        self.items_per_box: Optional[int] = None
        self.box_count_label: Optional[ttk.Label] = None
        self.total_qty_label: Optional[ttk.Label] = None
        self.allocation_text: Optional[tk.Text] = None
        self.allocation_result: Optional[AllocationResult] = None

        if worksheet is not None:
            self.total_qty = self._read_total_qty_from_excel(worksheet)
            self.items_per_box = self._extract_divisor_from_formula(worksheet)
            logger.info(f"Đã đọc thông tin từ Excel - Total QTY: {self.total_qty}, Items per box: {self.items_per_box}")

        self.dialog_config = DialogConfigManager()

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Nhap So Luong Cho Tung Size")

        width, height = self.dialog_config.get_dialog_size('size_quantity_input')
        self.dialog.geometry(f"{width}x{height}")
        self.dialog.resizable(True, True)

        self.dialog.transient(parent)
        self.dialog.grab_set()

        self.dialog.protocol("WM_DELETE_WINDOW", self._on_closing)

        self._create_widgets()
        self._center_window()

        logger.info(f"Đã mở dialog nhập số lượng cho {len(self.selected_sizes)} sizes")
    
    def _center_window(self) -> None:
        self.dialog.update_idletasks()

        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)

        self.dialog.geometry(f'{width}x{height}+{x}+{y}')

    def _read_total_qty_from_excel(self, worksheet) -> Optional[int]:
        """
        Đọc tổng số lượng (Tot QTY) từ hàng 15 hoặc 16 trong Excel.

        Args:
            worksheet: COM object của Excel worksheet

        Returns:
            Tổng số lượng (int) nếu tìm thấy, None nếu không tìm thấy hoặc có lỗi
        """
        try:
            for row in [15, 16]:
                used_range = worksheet.UsedRange
                last_col = used_range.Columns.Count

                for col in range(1, last_col + 1):
                    cell_value = worksheet.Cells(row, col).Value
                    if cell_value and isinstance(cell_value, str):
                        if "Tot QTY" in cell_value or "Total QTY" in cell_value:
                            total_value = worksheet.Cells(row, last_col).Value
                            if total_value is not None:
                                try:
                                    total_qty = int(total_value)
                                    logger.info(f"Đọc được Tot QTY từ hàng {row}: {total_qty}")
                                    return total_qty
                                except (ValueError, TypeError):
                                    logger.warning(f"Giá trị Tot QTY không hợp lệ ở hàng {row}: {total_value}")
                                    continue

            logger.warning("Không tìm thấy Tot QTY trong hàng 15/16")
            return None

        except Exception as e:
            logger.error(f"Lỗi khi đọc Tot QTY từ Excel: {e}", exc_info=True)
            return None

    def _extract_divisor_from_formula(self, worksheet) -> Optional[int]:
        """
        Extract số chia (items per box) từ công thức Excel ở ô G18.

        Công thức có dạng: =SUM(...)/20 hoặc =A1/20
        Method sẽ extract số 20 từ công thức.

        Args:
            worksheet: COM object của Excel worksheet

        Returns:
            Số chia (int) nếu parse thành công, None nếu không parse được hoặc có lỗi
        """
        try:
            formula = worksheet.Cells(18, 7).Formula

            if not formula or not isinstance(formula, str):
                logger.warning("Ô G18 không chứa công thức")
                return None

            match = re.search(r'/\s*(\d+)\s*$', formula)
            if match:
                divisor = int(match.group(1))
                logger.info(f"Extract được items per box từ công thức G18: {divisor}")
                return divisor
            else:
                logger.warning(f"Không parse được số chia từ công thức G18: {formula}")
                return None

        except Exception as e:
            logger.error(f"Lỗi khi extract divisor từ công thức G18: {e}", exc_info=True)
            return None

    def _calculate_box_count(self) -> int:
        """
        Tính số thùng cần đóng dựa trên tổng số lượng đã nhập và items per box.

        Công thức: box_count = ceil(tổng số lượng / items_per_box)

        Returns:
            Số thùng cần đóng (int), 0 nếu items_per_box = None/0 hoặc có lỗi
        """
        try:
            total_qty = 0
            for entry in self.quantity_entries.values():
                value = entry.get().strip()
                if value.isdigit():
                    total_qty += int(value)

            if not self.items_per_box or self.items_per_box == 0:
                return 0

            box_count = math.ceil(total_qty / self.items_per_box)
            return box_count

        except Exception as e:
            logger.error(f"Lỗi khi tính box count: {e}", exc_info=True)
            return 0

    def _update_box_count_display(self, event=None) -> None:
        try:
            size_quantities: Dict[str, int] = {}
            total_qty = 0

            for size, entry in self.quantity_entries.items():
                value = entry.get().strip()
                if value.isdigit():
                    qty = int(value)
                    total_qty += qty
                    size_quantities[size] = qty

            if self.total_qty_label:
                if total_qty > 0:
                    self.total_qty_label.config(
                        text=f"Tong so luong da nhap: {total_qty} cai",
                        foreground='blue'
                    )
                else:
                    self.total_qty_label.config(
                        text="Tong so luong da nhap: 0 cai",
                        foreground='gray'
                    )

            if not self.items_per_box or self.items_per_box == 0:
                return

            if size_quantities:
                calc = CartonAllocationCalculator(self.items_per_box)
                self.allocation_result = calc.get_full_result(size_quantities)
                self._update_allocation_display()

                if self.box_count_label:
                    total_boxes = self.allocation_result.total_boxes
                    if total_boxes > 0:
                        self.box_count_label.config(
                            text=f"Tong: {total_boxes} thung ({self.allocation_result.total_full_boxes} nguyen + {self.allocation_result.total_combined_boxes} ghep)",
                            foreground='green'
                        )
                    else:
                        self.box_count_label.config(
                            text="So thung can dong: 0 thung",
                            foreground='gray'
                        )
            else:
                self.allocation_result = None
                if self.allocation_text:
                    self.allocation_text.config(state=tk.NORMAL)
                    self.allocation_text.delete('1.0', tk.END)
                    self.allocation_text.config(state=tk.DISABLED)
                if self.box_count_label:
                    self.box_count_label.config(
                        text="So thung can dong: 0 thung",
                        foreground='gray'
                    )

        except Exception as e:
            logger.error(f"Loi khi update box count display: {e}", exc_info=True)

    def _update_allocation_display(self) -> None:
        if not self.allocation_text or not self.allocation_result:
            return

        self.allocation_text.config(state=tk.NORMAL)
        self.allocation_text.delete('1.0', tk.END)

        self.allocation_text.insert(tk.END, "=== CHI TIET PHAN BO ===\n\n")

        for size, alloc in self.allocation_result.allocations.items():
            if alloc.remainder > 0:
                line = f"  {size}: {alloc.total_pcs} pcs -> {alloc.full_boxes} thung ({alloc.full_qty}) + {alloc.remainder} du\n"
            else:
                line = f"  {size}: {alloc.total_pcs} pcs -> {alloc.full_boxes} thung ({alloc.full_qty})\n"
            self.allocation_text.insert(tk.END, line)

        if self.allocation_result.combined_cartons:
            self.allocation_text.insert(tk.END, "\n=== THUNG GHEP ===\n\n")

            for i, carton in enumerate(self.allocation_result.combined_cartons, 1):
                details = ' + '.join([f'{s}({q})' for s, q in carton.quantities.items()])
                full_marker = "[FULL]" if carton.is_full(self.items_per_box) else ""
                line = f"  Thung {i}: {details} = {carton.total_pcs} pcs {full_marker}\n"
                self.allocation_text.insert(tk.END, line)

        self.allocation_text.config(state=tk.DISABLED)

    def _on_canvas_mouse_wheel(self, event) -> None:
        """
        Xử lý sự kiện cuộn chuột để scroll canvas (Windows/macOS).

        Args:
            event: Mouse wheel event object
        """
        try:
            if not self.canvas:
                return

            if event.delta > 0:
                self.canvas.yview_scroll(-1, "units")
            elif event.delta < 0:
                self.canvas.yview_scroll(1, "units")

        except Exception as e:
            logger.error(f"Lỗi khi xử lý canvas mouse wheel: {e}", exc_info=True)

    def _on_canvas_mouse_wheel_linux(self, event, direction: int) -> None:
        """
        Xử lý sự kiện cuộn chuột để scroll canvas (Linux).

        Args:
            event: Mouse event object
            direction: 1 (scroll up) hoặc -1 (scroll down)
        """
        try:
            if not self.canvas:
                return

            self.canvas.yview_scroll(-direction, "units")

        except Exception as e:
            logger.error(f"Lỗi khi xử lý canvas mouse wheel Linux: {e}", exc_info=True)
    
    def _create_widgets(self) -> None:
        header_frame = ttk.Frame(self.dialog)
        header_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(
            header_frame,
            text="Nhập Số Lượng Cho Từng Size",
            font=('Arial', 12, 'bold')
        ).pack(anchor=tk.W)
        
        ttk.Label(
            header_frame,
            text="Nhập số lượng (1-1000) hoặc để trống",
            font=('Arial', 9),
            foreground='gray'
        ).pack(anchor=tk.W, pady=(5, 0))

        ttk.Label(
            header_frame,
            text="💡 Mẹo: Cuộn chuột để scroll danh sách",
            font=('Arial', 8),
            foreground='#0066cc'
        ).pack(anchor=tk.W, pady=(2, 0))
        
        scroll_frame = ttk.LabelFrame(self.dialog, text="Số Lượng Từng Size", padding=10)
        scroll_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.canvas = tk.Canvas(scroll_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        scrollable_frame = ttk.Frame(self.canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        for idx, size in enumerate(self.selected_sizes):
            row_frame = ttk.Frame(scrollable_frame)
            row_frame.pack(fill=tk.X, pady=5, padx=10)

            ttk.Label(
                row_frame,
                text=f"Size {size}:",
                width=15,
                anchor=tk.W
            ).pack(side=tk.LEFT, padx=(0, 10))

            entry = ttk.Entry(row_frame, width=15)
            entry.pack(side=tk.LEFT, padx=(0, 10))
            self.quantity_entries[size] = entry

            entry.bind('<KeyRelease>', self._update_box_count_display)

            if size in self.current_quantities and self.current_quantities[size] is not None:
                entry.insert(0, str(self.current_quantities[size]))

            ttk.Label(
                row_frame,
                text="thùng",
                foreground='gray'
            ).pack(side=tk.LEFT)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.bind('<MouseWheel>', self._on_canvas_mouse_wheel)
        self.canvas.bind('<Button-4>', lambda e: self._on_canvas_mouse_wheel_linux(e, 1))
        self.canvas.bind('<Button-5>', lambda e: self._on_canvas_mouse_wheel_linux(e, -1))

        if self.items_per_box is not None:
            box_count_frame = ttk.LabelFrame(self.dialog, text="Thong Tin Dong Goi", padding=10)
            box_count_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

            ttk.Label(
                box_count_frame,
                text=f"So luong moi thung: {self.items_per_box}",
                foreground='gray'
            ).pack(anchor=tk.W)

            self.total_qty_label = ttk.Label(
                box_count_frame,
                text="Tong so luong da nhap: 0 cai",
                foreground='gray'
            )
            self.total_qty_label.pack(anchor=tk.W, pady=(5, 0))

            self.box_count_label = ttk.Label(
                box_count_frame,
                text="So thung can dong: 0 thung",
                font=('Arial', 10, 'bold'),
                foreground='gray'
            )
            self.box_count_label.pack(anchor=tk.W, pady=(5, 0))

            alloc_frame = ttk.LabelFrame(self.dialog, text="Chi Tiet Phan Bo", padding=5)
            alloc_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=(0, 10))

            self.allocation_text = tk.Text(
                alloc_frame,
                height=8,
                width=50,
                font=('Consolas', 9),
                state=tk.DISABLED,
                wrap=tk.WORD
            )
            alloc_scrollbar = ttk.Scrollbar(alloc_frame, orient=tk.VERTICAL, command=self.allocation_text.yview)
            self.allocation_text.configure(yscrollcommand=alloc_scrollbar.set)
            self.allocation_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            alloc_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            self._update_box_count_display()

        action_frame = ttk.Frame(self.dialog)
        action_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        ttk.Button(
            action_frame,
            text="Áp dụng",
            command=self._save_quantities,
            width=15
        ).pack(side=tk.RIGHT, padx=(5, 0))
        
        ttk.Button(
            action_frame,
            text="Hủy",
            command=self.dialog.destroy,
            width=15
        ).pack(side=tk.RIGHT)
    
    def _validate_quantities(self) -> bool:
        for size, entry in self.quantity_entries.items():
            value = entry.get().strip()
            
            if value == "":
                continue
            
            try:
                quantity = int(value)
                if quantity < 1 or quantity > 1000:
                    messagebox.showerror(
                        "Lỗi Validation",
                        f"Size {size}: Số lượng phải từ 1 đến 1000!\n\n"
                        f"Giá trị nhập: {quantity}"
                    )
                    entry.focus_set()
                    return False
            except ValueError:
                messagebox.showerror(
                    "Lỗi Validation",
                    f"Size {size}: Số lượng phải là số nguyên!\n\n"
                    f"Giá trị nhập: '{value}'"
                )
                entry.focus_set()
                return False
        
        return True
    
    def _save_quantities(self) -> None:
        if not self._validate_quantities():
            return

        self.quantities.clear()

        for size, entry in self.quantity_entries.items():
            value = entry.get().strip()

            if value != "":
                self.quantities[size] = int(value)
            elif size in self.current_quantities and self.current_quantities[size] is not None:
                self.quantities[size] = None

        logger.info(f"Đã nhập số lượng cho {len(self.quantities)} sizes")
        self._save_size_and_close()

    def _on_closing(self) -> None:
        self._save_size_and_close()

    def _save_size_and_close(self) -> None:
        try:
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            self.dialog_config.save_dialog_size('size_quantity_input', width, height)
        except Exception:
            pass
        finally:
            self.dialog.destroy()

    def show(self) -> None:
        self.parent.wait_window(self.dialog)

    def get_quantities(self) -> Dict[str, int]:
        return self.quantities

    def get_allocation_result(self) -> Optional[AllocationResult]:
        return self.allocation_result

    def get_items_per_box(self) -> Optional[int]:
        return self.items_per_box
