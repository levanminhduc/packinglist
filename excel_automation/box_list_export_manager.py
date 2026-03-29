from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional
from win32com.client import CDispatch
import win32clipboard
import logging

from excel_automation.box_list_export_config import BoxListExportConfig
from excel_automation.utils import get_size_sort_key, normalize_size_value

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

        size_column_number = self._column_letter_to_number(size_column)

        try:
            used_range = worksheet.UsedRange
            max_column = used_range.Column + used_range.Columns.Count - 1
            scan_end_column = max(39, max_column + 1)
        except Exception:
            scan_end_column = 39

        size_data_end_row = self.config.get_size_data_end_row()
        try:
            range_str = f"{size_column}{size_data_start_row}:{size_column}{size_data_end_row}"
            raw_size_values = worksheet.Range(range_str).Value
        except Exception:
            raw_size_values = None

        size_to_row: Dict[str, int] = {}
        if raw_size_values is not None:
            if not isinstance(raw_size_values, tuple):
                raw_size_values = ((raw_size_values,),)
            for row_offset, row_tuple in enumerate(raw_size_values):
                cell_value = row_tuple[0] if isinstance(row_tuple, tuple) else row_tuple
                if cell_value is not None and str(cell_value).strip() != "":
                    size_str = normalize_size_value(cell_value)
                    if size_str:
                        size_to_row[size_str] = size_data_start_row + row_offset

        start_col = 7
        end_col = scan_end_column - 1

        try:
            box_start_values_raw = worksheet.Range(
                worksheet.Cells(box_start_row, start_col),
                worksheet.Cells(box_start_row, end_col)
            ).Value
        except Exception:
            box_start_values_raw = None

        try:
            box_end_values_raw = worksheet.Range(
                worksheet.Cells(box_end_row, start_col),
                worksheet.Cells(box_end_row, end_col)
            ).Value
        except Exception:
            box_end_values_raw = None

        if box_start_values_raw is not None and not isinstance(box_start_values_raw, tuple):
            box_start_values_raw = ((box_start_values_raw,),)
        if box_end_values_raw is not None and not isinstance(box_end_values_raw, tuple):
            box_end_values_raw = ((box_end_values_raw,),)

        box_start_values = box_start_values_raw[0] if box_start_values_raw else ()
        box_end_values = box_end_values_raw[0] if box_end_values_raw else ()

        box_ranges: Dict[str, List[Tuple[int, int, int, int]]] = {}

        for size in selected_sizes:
            if size not in size_to_row:
                logger.warning(f"Size {size} không tìm thấy trong cột {size_column}")
                box_ranges[size] = []
                continue

            size_row = size_to_row[size]

            try:
                size_row_values_raw = worksheet.Range(
                    worksheet.Cells(size_row, start_col),
                    worksheet.Cells(size_row, end_col)
                ).Value
            except Exception:
                box_ranges[size] = []
                continue

            if size_row_values_raw is not None and not isinstance(size_row_values_raw, tuple):
                size_row_values_raw = ((size_row_values_raw,),)

            size_row_values = size_row_values_raw[0] if size_row_values_raw else ()

            size_box_ranges: List[Tuple[int, int, int, int]] = []

            for col_offset in range(len(size_row_values)):
                column_number = start_col + col_offset
                quantity_value = size_row_values[col_offset]

                if quantity_value is None:
                    continue

                try:
                    quantity = int(float(quantity_value))
                    if quantity <= 0:
                        continue
                except (ValueError, TypeError):
                    continue

                if col_offset >= len(box_start_values) or col_offset >= len(box_end_values):
                    continue

                box_start_value = box_start_values[col_offset]
                box_end_value = box_end_values[col_offset]

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

            new_sheet.Cells(start_row, start_col_num).Value = header
            new_sheet.Cells(start_row, start_col_num).Font.Bold = True
            new_sheet.Cells(start_row, start_col_num).Font.Size = 20
            new_sheet.Cells(start_row, start_col_num).HorizontalAlignment = -4131

            for col_idx, column_lines in enumerate(columns):
                col_num = start_col_num + col_idx
                data_start_row = start_row + header_rows

                if column_lines:
                    data_array = [[line] for line in column_lines]
                    data_end_row = data_start_row + len(column_lines) - 1
                    new_sheet.Range(
                        new_sheet.Cells(data_start_row, col_num),
                        new_sheet.Cells(data_end_row, col_num)
                    ).Value = data_array

                    new_sheet.Range(
                        new_sheet.Cells(data_start_row, col_num),
                        new_sheet.Cells(data_end_row, col_num)
                    ).HorizontalAlignment = -4108

                    bold_rows = []
                    non_bold_rows = []
                    for line_idx, line in enumerate(column_lines):
                        row_num = data_start_row + line_idx
                        if line.startswith("SIZE "):
                            bold_rows.append(row_num)
                        else:
                            non_bold_rows.append(row_num)

                    if bold_rows:
                        bold_range = self._build_union_range(
                            new_sheet, bold_rows, col_num
                        )
                        if bold_range:
                            bold_range.Font.Bold = True

                    if non_bold_rows:
                        non_bold_range = self._build_union_range(
                            new_sheet, non_bold_rows, col_num
                        )
                        if non_bold_range:
                            non_bold_range.Font.Bold = False

            logger.info(f"Đã paste và format {len(columns)} cột vào sheet mới: {new_sheet.Name}")
            return True
        except Exception as e:
            logger.error(f"Lỗi khi paste vào Excel: {e}", exc_info=True)
            return False
    
    def step_read_box_ranges(
        self,
        worksheet: CDispatch,
        selected_sizes: List[str]
    ) -> Dict[str, List[Tuple[int, int, int, int]]]:
        return self.read_box_ranges(worksheet, selected_sizes)

    def step_analyze_and_build_result(
        self,
        workbook: CDispatch,
        worksheet: CDispatch,
        selected_sizes: List[str],
        box_ranges_dict: Dict[str, List[Tuple[int, int, int, int]]],
        items_per_box: Optional[int] = None
    ) -> BoxListExportResult:
        valid_count = sum(
            1 for size_ranges in box_ranges_dict.values()
            if len(size_ranges) > 0
        )

        if valid_count == 0:
            return BoxListExportResult(
                success=False,
                error_message="Không có size nào có dữ liệu box hợp lệ"
            )

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

        return BoxListExportResult(
            success=True,
            text=text,
            box_ranges=box_ranges,
            total_boxes=total_boxes,
            header=header,
            total_columns=total_columns
        )

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

            box_ranges_dict = self.step_read_box_ranges(worksheet, selected_sizes)

            result = self.step_analyze_and_build_result(
                workbook, worksheet, selected_sizes,
                box_ranges_dict, items_per_box
            )

            if result.success:
                self.copy_to_clipboard(result.text)
                logger.info(
                    f"Hoàn thành xuất danh sách thùng: "
                    f"{result.total_boxes} thùng, {result.total_columns} cột"
                )

            return result

        except Exception as e:
            error_msg = f"Lỗi khi xuất danh sách thùng: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return BoxListExportResult(success=False, error_message=error_msg)
        finally:
            excel_app.ScreenUpdating = True
    
    def _build_union_range(
        self,
        sheet: CDispatch,
        rows: List[int],
        col_num: int
    ) -> Optional[CDispatch]:
        if not rows:
            return None

        try:
            excel_app = sheet.Application
            result_range = sheet.Cells(rows[0], col_num)

            batch_size = 30
            for i in range(1, len(rows), batch_size):
                batch = rows[i:i + batch_size]
                for row in batch:
                    result_range = excel_app.Union(
                        result_range, sheet.Cells(row, col_num)
                    )

            return result_range
        except Exception as e:
            logger.warning(f"Không thể tạo union range, fallback từng cell: {e}")
            return None

    def _column_letter_to_number(self, column: str) -> int:
        column = column.upper()
        result = 0
        for char in column:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

