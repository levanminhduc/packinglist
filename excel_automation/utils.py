"""
Module chứa các hàm tiện ích chung.
"""

import os
import shutil
from pathlib import Path
from datetime import datetime
from typing import List, Optional
import logging

logger = logging.getLogger(__name__)


def setup_logging(
    log_file: Optional[str] = None,
    level: int = logging.INFO
) -> None:
    """
    Cấu hình logging.
    
    Args:
        log_file: Đường dẫn file log
        level: Mức độ log
    """
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    handlers = [logging.StreamHandler()]
    
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding='utf-8'))
    
    logging.basicConfig(
        level=level,
        format=log_format,
        handlers=handlers
    )
    
    logger.info("Logging đã được cấu hình")


def get_size_sort_key(size: str) -> tuple:
    """
    Tạo sort key cho size để sắp xếp theo thứ tự: XS, S, M, L, XL, XXL, XXXL.

    Args:
        size: Size cần sắp xếp

    Returns:
        Tuple (priority, numeric_value) để sắp xếp
    """
    size_upper = size.upper().strip()

    size_order = {
        'XS': 0,
        'S': 1,
        'M': 2,
        'L': 3,
        'XL': 4,
        'XXL': 5,
        'XXXL': 6
    }

    if size_upper in size_order:
        return (0, size_order[size_upper])

    if size.replace('.', '').replace('-', '').isdigit():
        try:
            return (1, float(size))
        except ValueError:
            return (2, size)

    return (2, size)


def get_timestamp(format_string: str = "%Y%m%d_%H%M%S") -> str:
    """
    Lấy timestamp hiện tại.

    Args:
        format_string: Format của timestamp

    Returns:
        Timestamp string
    """
    return datetime.now().strftime(format_string)


def create_backup(
    file_path: str,
    backup_dir: str = "data/backup"
) -> str:
    """
    Tạo backup của file.
    
    Args:
        file_path: Đường dẫn file cần backup
        backup_dir: Thư mục lưu backup
        
    Returns:
        Đường dẫn file backup
    """
    try:
        source = Path(file_path)
        if not source.exists():
            raise FileNotFoundError(f"File không tồn tại: {file_path}")
        
        backup_path = Path(backup_dir)
        backup_path.mkdir(parents=True, exist_ok=True)
        
        timestamp = get_timestamp()
        backup_file = backup_path / f"{source.stem}_{timestamp}{source.suffix}"
        
        shutil.copy2(source, backup_file)
        logger.info(f"Backup file: {backup_file}")
        
        return str(backup_file)
    except Exception as e:
        logger.error(f"Lỗi khi tạo backup: {e}")
        raise


def list_excel_files(directory: str) -> List[str]:
    """
    Liệt kê tất cả file Excel trong thư mục.
    
    Args:
        directory: Đường dẫn thư mục
        
    Returns:
        Danh sách đường dẫn file Excel
    """
    try:
        dir_path = Path(directory)
        if not dir_path.exists():
            raise FileNotFoundError(f"Thư mục không tồn tại: {directory}")
        
        excel_extensions = ['.xlsx', '.xls', '.xlsm', '.xlsb']
        excel_files = []
        
        for ext in excel_extensions:
            excel_files.extend([str(f) for f in dir_path.glob(f'*{ext}')])
        
        logger.info(f"Tìm thấy {len(excel_files)} file Excel trong {directory}")
        return sorted(excel_files)
    except Exception as e:
        logger.error(f"Lỗi khi liệt kê file: {e}")
        raise


def validate_file_path(file_path: str, must_exist: bool = False) -> bool:
    """
    Kiểm tra tính hợp lệ của đường dẫn file.
    
    Args:
        file_path: Đường dẫn file
        must_exist: File phải tồn tại hay không
        
    Returns:
        True nếu hợp lệ
    """
    try:
        path = Path(file_path)
        
        if must_exist and not path.exists():
            logger.error(f"File không tồn tại: {file_path}")
            return False
        
        if path.suffix.lower() not in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
            logger.error(f"Không phải file Excel: {file_path}")
            return False
        
        return True
    except Exception as e:
        logger.error(f"Lỗi khi validate file path: {e}")
        return False


def ensure_directory(directory: str) -> None:
    """
    Đảm bảo thư mục tồn tại, tạo mới nếu chưa có.
    
    Args:
        directory: Đường dẫn thư mục
    """
    try:
        Path(directory).mkdir(parents=True, exist_ok=True)
        logger.info(f"Thư mục đã sẵn sàng: {directory}")
    except Exception as e:
        logger.error(f"Lỗi khi tạo thư mục: {e}")
        raise


def get_file_size(file_path: str) -> str:
    """
    Lấy kích thước file dạng human-readable.
    
    Args:
        file_path: Đường dẫn file
        
    Returns:
        Kích thước file (vd: "1.5 MB")
    """
    try:
        size_bytes = Path(file_path).stat().st_size
        
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        
        return f"{size_bytes:.2f} TB"
    except Exception as e:
        logger.error(f"Lỗi khi lấy file size: {e}")
        return "Unknown"


def clean_old_backups(
    backup_dir: str = "data/backup",
    keep_days: int = 30
) -> None:
    """
    Xóa các file backup cũ.
    
    Args:
        backup_dir: Thư mục backup
        keep_days: Số ngày giữ lại
    """
    try:
        backup_path = Path(backup_dir)
        if not backup_path.exists():
            return
        
        cutoff_time = datetime.now().timestamp() - (keep_days * 24 * 60 * 60)
        deleted_count = 0
        
        for file in backup_path.glob('*'):
            if file.is_file() and file.stat().st_mtime < cutoff_time:
                file.unlink()
                deleted_count += 1
        
        logger.info(f"Đã xóa {deleted_count} file backup cũ")
    except Exception as e:
        logger.error(f"Lỗi khi xóa backup cũ: {e}")
        raise


def convert_column_letter_to_index(column_letter: str) -> int:
    """
    Chuyển đổi chữ cái cột thành index số.
    
    Args:
        column_letter: Chữ cái cột (vd: 'A', 'AB')
        
    Returns:
        Index số (1-based)
    """
    result = 0
    for char in column_letter.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def convert_index_to_column_letter(index: int) -> str:
    """
    Chuyển đổi index số thành chữ cái cột.
    
    Args:
        index: Index số (1-based)
        
    Returns:
        Chữ cái cột
    """
    result = ""
    while index > 0:
        index -= 1
        result = chr(index % 26 + ord('A')) + result
        index //= 26
    return result

