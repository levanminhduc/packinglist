from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional
from win32com.client import CDispatch
import win32clipboard
import logging

from excel_automation.box_list_export_config import BoxListExportConfig
from excel_automation.utils import get_size_sort_key

logger = logging.getLogger(__name__)


@dataclass
class BoxRange:
    sizes: List[str]
    box_start: int
    box_end: int
    column_number: int
    total_pcs: Optional[int] = None
    items_per_box: Optional[int] = None

    def is_valid(self) -> bool:
        return self.box_start <= self.box_end and self.box_start > 0

    def is_combined(self) -> bool:
        return len(self.sizes) > 1

    def is_partial(self) -> bool:
        if self.is_combined():
            return False
        if self.total_pcs is None or self.items_per_box is None:
            return False
        return self.total_pcs < self.items_per_box

    def get_size_label(self, separator: str = "/") -> str:
        formatted_sizes = [self._format_size(s) for s in self.sizes]
        label = separator.join(formatted_sizes)

        if self.is_partial() and self.total_pcs is not None:
            label = f"{label}/{self.total_pcs}PCS"

        return label

    def _format_size(self, size: str) -> str:
        try:
            if '.' in size:
                num = float(size)
                if num == int(num):
                    return str(int(num))
                return size

            num = int(size)
            return str(num)
        except (ValueError, TypeError):
            return size

    def get_box_numbers(self) -> List[int]:
        return list(range(self.box_start, self.box_end + 1))


@dataclass
class BoxListExportResult:
    success: bool
    text: str = ""
    error_message: str = ""
    box_ranges: List[BoxRange] = field(default_factory=list)
    total_boxes: int = 0
    header: str = ""
    total_columns: int = 1

    def get_summary(self) -> str:
        if not self.success:
            return f"Lỗi: {self.error_message}"

        combined_count = sum(1 for br in self.box_ranges if br.is_combined())
        single_count = len(self.box_ranges) - combined_count

        summary_parts = []
        if single_count > 0:
            summary_parts.append(f"{single_count} size đơn")
        if combined_count > 0:
            summary_parts.append(f"{combined_count} size kết hợp")

        summary = " và ".join(summary_parts) if summary_parts else "0 size"
        column_info = f", {self.total_columns} cột" if self.total_columns > 1 else ""
        return f"Đã xuất {summary}, tổng {self.total_boxes} thùng{column_info}"


class BoxListExportManager:
    
    def __init__(self, config: BoxListExportConfig):
        self.config = config
    
    def read_box_ranges(
        self,
        worksheet: CDispatch,
        selected_sizes: List[str]
    ) -> Dict[str, List[Tuple[int, int, int, int]]]:
        box_start_row = self.config.get_box_start_row()
        box_end_row = self.config.get_box_end_row()
        size_column = self.config.get_size_column()
        size_data_start_row = self.config.get_size_data_start_row()
        size_data_end_row = self.config.get_size_data_end_row()

        size_column_number = self._column_letter_to_number(size_column)

        size_to_row: Dict[str, int] = {}
        for row in range(size_data_start_row, size_data_end_row + 1):
            cell_value = worksheet.Cells(row, size_column_number).Value
            if cell_value is not None:
                size_str = str(cell_value).strip()
                if size_str.isdigit():
                    size_str = size_str.zfill(3)
                if size_str:
                    size_to_row[size_str] = row

        box_ranges: Dict[str, List[Tuple[int, int, int, int]]] = {}

        for size in selected_sizes:
            if size not in size_to_row:
                logger.warning(f"Size {size} không tìm thấy trong cột {size_column}")
                box_ranges[size] = []
                continue

            size_row = size_to_row[size]
            size_box_ranges: List[Tuple[int, int, int, int]] = []

            for column_number in range(7, 39):
                try:
                    quantity_value = worksheet.Cells(size_row, column_number).Value

                    if quantity_value is None:
                        continue

                    try:
                        quantity = int(float(quantity_value))
                        if quantity <= 0:
                            continue
                    except (ValueError, TypeError):
                        continue

                    box_start_value = worksheet.Cells(box_start_row, column_number).Value
                    box_end_value = worksheet.Cells(box_end_row, column_number).Value

                    if box_start_value is None or box_end_value is None:
                        continue

                    try:
                        box_start = int(box_start_value)
                        box_end = int(box_end_value)
                    except (ValueError, TypeError):
                        logger.warning(
                            f"Size {size}, cột {column_number}: box_start hoặc box_end không hợp lệ"
                        )
                        continue

                    if box_start > box_end:
                        logger.warning(
                            f"Size {size}, cột {column_number}: box_start ({box_start}) > box_end ({box_end})"
                        )
                        continue

                    size_box_ranges.append((box_start, box_end, column_number, quantity))

                except Exception as e:
                    logger.error(
                        f"Lỗi khi đọc box range cho size {size}, cột {column_number}: {e}",
                        exc_info=True
                    )
                    continue

            box_ranges[size] = size_box_ranges

        return box_ranges
    
    def detect_combined_sizes(
        self,
        selected_sizes: List[str],
        box_ranges: Dict[str, List[Tuple[int, int, int, int]]],
        items_per_box: Optional[int] = None
    ) -> List[BoxRange]:
        if not self.config.is_combined_detection_enabled():
            result = []
            for size in selected_sizes:
                size_ranges = box_ranges.get(size, [])
                for box_start, box_end, column_number, quantity in size_ranges:
                    result.append(BoxRange(
                        [size], box_start, box_end, column_number,
                        total_pcs=quantity, items_per_box=items_per_box
                    ))
            return result

        groups: Dict[Tuple[int, int], List[Tuple[str, int, int]]] = {}

        for size in selected_sizes:
            size_ranges = box_ranges.get(size, [])
            for box_start, box_end, column_number, quantity in size_ranges:
                key = (box_start, box_end)
                if key not in groups:
                    groups[key] = []
                groups[key].append((size, column_number, quantity))

        result = []
        for (box_start, box_end), size_list in groups.items():
            sizes_set = set(s for s, _, _ in size_list)
            sizes = list(sizes_set)
            column_number = size_list[0][1]
            total_pcs = sum(qty for _, _, qty in size_list)

            if self.config.is_sort_combined_sizes_enabled():
                sizes.sort(key=get_size_sort_key)

            result.append(BoxRange(
                sizes, box_start, box_end, column_number,
                total_pcs=total_pcs, items_per_box=items_per_box
            ))

        result = sorted(result, key=lambda br: (br.is_partial(), br.box_start))

        return result

    def get_filename(self, workbook: CDispatch) -> str:
        try:
            full_name = workbook.Name
            if '.' in full_name:
                return full_name.rsplit('.', 1)[0]
            return full_name
        except Exception as e:
            logger.warning(f"Không thể lấy tên file: {e}")
            return "Unknown"

    def get_po_number(self, worksheet: CDispatch) -> str:
        try:
            po_row = self.config.get_po_cell_row()
            po_col = self.config.get_po_cell_column()
            col_num = self._column_letter_to_number(po_col)

            cell_value = worksheet.Cells(po_row, col_num).Value
            if cell_value is not None:
                value_str = str(cell_value)
                if value_str.endswith('.0'):
                    return value_str[:-2]
                return value_str
            return ""
        except Exception as e:
            logger.warning(f"Không thể đọc PO number: {e}")
            return ""

    def generate_header(
        self,
        workbook: CDispatch,
        worksheet: CDispatch,
        items_per_box: Optional[int] = None
    ) -> str:
        filename = self.get_filename(workbook)
        po_number = self.get_po_number(worksheet)
        header = f"{filename}_PO:{po_number}"
        if items_per_box is not None:
            header = f"{header} / {items_per_box} PCS"
        return header

    def generate_sheet_name(self, workbook: CDispatch, worksheet: CDispatch) -> str:
        filename = self.get_filename(workbook)
        po_number = self.get_po_number(worksheet)

        if len(po_number) >= 4:
            po_suffix = po_number[-4:]
        else:
            po_suffix = po_number

        return f"{filename}_{po_suffix}"

    def create_new_sheet(self, workbook: CDispatch, worksheet: CDispatch) -> CDispatch:
        try:
            sheet_name = self.generate_sheet_name(workbook, worksheet)

            existing_names = [sheet.Name for sheet in workbook.Worksheets]
            if sheet_name in existing_names:
                counter = 1
                while f"{sheet_name}_{counter}" in existing_names:
                    counter += 1
                sheet_name = f"{sheet_name}_{counter}"

            new_sheet = workbook.Worksheets.Add()
            new_sheet.Name = sheet_name

            logger.info(f"Đã tạo sheet mới: {sheet_name}")
            return new_sheet
        except Exception as e:
            logger.error(f"Lỗi khi tạo sheet mới: {e}", exc_info=True)
            raise

    def split_into_columns(self, lines: List[str]) -> List[List[str]]:
        max_rows = self.config.get_max_rows_per_column()
        header_rows = self.config.get_header_rows()
        max_content_rows = max_rows - header_rows

        columns = []
        current_column = []

        for line in lines:
            if len(current_column) >= max_content_rows:
                columns.append(current_column)
                current_column = []
            current_column.append(line)

        if current_column:
            columns.append(current_column)

        return columns

    def generate_box_list_text(self, box_ranges: List[BoxRange]) -> str:
        separator = self.config.get_combined_size_separator()
        lines = []

        for box_range in box_ranges:
            size_label = box_range.get_size_label(separator)
            lines.append(f"SIZE {size_label}")

            for box_number in box_range.get_box_numbers():
                lines.append(str(box_number))

        return "\n".join(lines)
    
    def copy_to_clipboard(self, text: str) -> bool:
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            logger.info("Đã copy text vào clipboard")
            return True
        except Exception as e:
            logger.error(f"Lỗi khi copy vào clipboard: {e}", exc_info=True)
            return False

    def paste_and_format_to_excel(
        self,
        workbook: CDispatch,
        worksheet: CDispatch,
        box_ranges: List[BoxRange],
        new_sheet: CDispatch,
        start_column: str = "A",
        start_row: int = 1,
        items_per_box: Optional[int] = None
    ) -> bool:
        try:
            header = self.generate_header(workbook, worksheet, items_per_box)
            header_rows = self.config.get_header_rows()

            content_lines = []
            for box_range in box_ranges:
                separator = self.config.get_combined_size_separator()
                size_label = box_range.get_size_label(separator)
                content_lines.append(f"SIZE {size_label}")

                for box_number in box_range.get_box_numbers():
                    content_lines.append(str(box_number))

            columns = self.split_into_columns(content_lines)
            start_col_num = self._column_letter_to_number(start_column)

            for col_idx, column_lines in enumerate(columns):
                col_num = start_col_num + col_idx
                current_row = start_row

                if col_idx == 0:
                    new_sheet.Cells(current_row, col_num).Value = header
                    new_sheet.Cells(current_row, col_num).Font.Bold = True
                    new_sheet.Cells(current_row, col_num).Font.Size = 20
                    new_sheet.Cells(current_row, col_num).HorizontalAlignment = -4131

                current_row += header_rows

                for line in column_lines:
                    new_sheet.Cells(current_row, col_num).Value = line
                    new_sheet.Cells(current_row, col_num).HorizontalAlignment = -4108

                    if line.startswith("SIZE "):
                        new_sheet.Cells(current_row, col_num).Font.Bold = True
                    else:
                        new_sheet.Cells(current_row, col_num).Font.Bold = False

                    current_row += 1

            logger.info(f"Đã paste và format {len(columns)} cột vào sheet mới: {new_sheet.Name}")
            return True
        except Exception as e:
            logger.error(f"Lỗi khi paste vào Excel: {e}", exc_info=True)
            return False
    
    def export_box_list(
        self,
        excel_app: CDispatch,
        workbook: CDispatch,
        worksheet: CDispatch,
        selected_sizes: List[str],
        items_per_box: Optional[int] = None
    ) -> BoxListExportResult:
        logger.info(f"Bắt đầu xuất danh sách thùng cho {len(selected_sizes)} sizes")

        try:
            excel_app.ScreenUpdating = False

            box_ranges_dict = self.read_box_ranges(worksheet, selected_sizes)

            valid_count = sum(
                1 for size_ranges in box_ranges_dict.values()
                if len(size_ranges) > 0
            )

            if valid_count == 0:
                error_msg = "Không có size nào có dữ liệu box hợp lệ"
                logger.warning(error_msg)
                return BoxListExportResult(success=False, error_message=error_msg)

            box_ranges = self.detect_combined_sizes(
                selected_sizes, box_ranges_dict, items_per_box
            )

            partial_count = sum(1 for br in box_ranges if br.is_partial())
            logger.info(
                f"Đã phát hiện {len(box_ranges)} box ranges "
                f"({sum(1 for br in box_ranges if br.is_combined())} kết hợp, "
                f"{partial_count} thùng lẻ)"
            )

            header = self.generate_header(workbook, worksheet, items_per_box)

            content_lines = []
            for box_range in box_ranges:
                separator = self.config.get_combined_size_separator()
                size_label = box_range.get_size_label(separator)
                content_lines.append(f"SIZE {size_label}")

                for box_number in box_range.get_box_numbers():
                    content_lines.append(str(box_number))

            columns = self.split_into_columns(content_lines)
            total_columns = len(columns)

            text_parts = []
            for col_idx, column_lines in enumerate(columns):
                text_parts.append(header)
                text_parts.append("")
                text_parts.extend(column_lines)
                if col_idx < len(columns) - 1:
                    text_parts.append("")

            text = "\n".join(text_parts)

            total_boxes = sum(len(br.get_box_numbers()) for br in box_ranges)

            clipboard_success = self.copy_to_clipboard(text)

            if not clipboard_success:
                logger.warning("Không thể copy vào clipboard, nhưng vẫn trả về kết quả")

            logger.info(f"Hoàn thành xuất danh sách thùng: {total_boxes} thùng, {total_columns} cột")

            return BoxListExportResult(
                success=True,
                text=text,
                box_ranges=box_ranges,
                total_boxes=total_boxes,
                header=header,
                total_columns=total_columns
            )
            
        except Exception as e:
            error_msg = f"Lỗi khi xuất danh sách thùng: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return BoxListExportResult(success=False, error_message=error_msg)
        finally:
            excel_app.ScreenUpdating = True
    
    def _column_letter_to_number(self, column: str) -> int:
        column = column.upper()
        result = 0
        for char in column:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

