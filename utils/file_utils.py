"""
File Utilities
Các hàm tiện ích cho xử lý file và thư mục
Author: Assistant
Date: 2025-07-26
"""

import os
import shutil
from pathlib import Path
from typing import List, Optional
from datetime import datetime


def ensure_directory(dir_path: str) -> bool:
    """
    Đảm bảo thư mục tồn tại, tạo nếu chưa có
    
    Args:
        dir_path (str): Đường dẫn thư mục
        
    Returns:
        bool: True nếu thành công
    """
    try:
        Path(dir_path).mkdir(parents=True, exist_ok=True)
        return True
    except Exception as e:
        print(f"❌ Lỗi tạo thư mục {dir_path}: {e}")
        return False


def ensure_directories(dir_paths: List[str]) -> bool:
    """
    Đảm bảo nhiều thư mục tồn tại
    
    Args:
        dir_paths (List[str]): Danh sách đường dẫn thư mục
        
    Returns:
        bool: True nếu tất cả thành công
    """
    success = True
    for dir_path in dir_paths:
        if not ensure_directory(dir_path):
            success = False
    return success


def get_file_timestamp(file_path: str) -> Optional[datetime]:
    """
    Lấy timestamp của file
    
    Args:
        file_path (str): Đường dẫn file
        
    Returns:
        Optional[datetime]: Timestamp hoặc None nếu lỗi
    """
    try:
        path = Path(file_path)
        if path.exists():
            return datetime.fromtimestamp(path.stat().st_mtime)
        return None
    except Exception:
        return None


def get_file_size(file_path: str) -> int:
    """
    Lấy kích thước file
    
    Args:
        file_path (str): Đường dẫn file
        
    Returns:
        int: Kích thước file (bytes), 0 nếu lỗi
    """
    try:
        return Path(file_path).stat().st_size
    except Exception:
        return 0


def format_file_size(size_bytes: int) -> str:
    """
    Format kích thước file dễ đọc
    
    Args:
        size_bytes (int): Kích thước tính bằng bytes
        
    Returns:
        str: Kích thước đã format
    """
    if size_bytes == 0:
        return "0 B"
    
    size_names = ["B", "KB", "MB", "GB", "TB"]
    i = 0
    while size_bytes >= 1024 and i < len(size_names) - 1:
        size_bytes /= 1024.0
        i += 1
    
    return f"{size_bytes:.1f} {size_names[i]}"


def list_files_with_pattern(directory: str, pattern: str = "*") -> List[str]:
    """
    Liệt kê files theo pattern
    
    Args:
        directory (str): Thư mục cần tìm
        pattern (str): Pattern tìm kiếm (VD: "*.xlsx")
        
    Returns:
        List[str]: Danh sách file paths
    """
    try:
        path = Path(directory)
        if not path.exists():
            return []
        
        return [str(f) for f in path.glob(pattern) if f.is_file()]
    except Exception:
        return []


def get_latest_file(directory: str, pattern: str = "*") -> Optional[str]:
    """
    Lấy file mới nhất trong thư mục
    
    Args:
        directory (str): Thư mục cần tìm
        pattern (str): Pattern tìm kiếm
        
    Returns:
        Optional[str]: Đường dẫn file mới nhất hoặc None
    """
    files = list_files_with_pattern(directory, pattern)
    if not files:
        return None
    
    # Sắp xếp theo thời gian modification
    files_with_time = [(f, get_file_timestamp(f)) for f in files]
    files_with_time = [(f, t) for f, t in files_with_time if t is not None]
    
    if not files_with_time:
        return None
    
    files_with_time.sort(key=lambda x: x[1], reverse=True)
    return files_with_time[0][0]


def backup_file(file_path: str, backup_dir: str = "backups") -> Optional[str]:
    """
    Sao lưu file
    
    Args:
        file_path (str): Đường dẫn file cần backup
        backup_dir (str): Thư mục backup
        
    Returns:
        Optional[str]: Đường dẫn file backup hoặc None nếu lỗi
    """
    try:
        source_path = Path(file_path)
        if not source_path.exists():
            return None
        
        # Tạo thư mục backup
        backup_path = Path(backup_dir)
        backup_path.mkdir(exist_ok=True)
        
        # Tạo tên file backup với timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"{source_path.stem}_{timestamp}{source_path.suffix}"
        backup_file_path = backup_path / backup_filename
        
        # Copy file
        shutil.copy2(source_path, backup_file_path)
        return str(backup_file_path)
        
    except Exception as e:
        print(f"❌ Lỗi backup file: {e}")
        return None


def clean_old_files(directory: str, pattern: str = "*", keep_count: int = 5) -> int:
    """
    Xóa các file cũ, chỉ giữ lại số lượng nhất định
    
    Args:
        directory (str): Thư mục cần dọn dẹp
        pattern (str): Pattern file cần dọn
        keep_count (int): Số file cần giữ lại
        
    Returns:
        int: Số file đã xóa
    """
    try:
        files = list_files_with_pattern(directory, pattern)
        if len(files) <= keep_count:
            return 0
        
        # Sắp xếp theo thời gian modification
        files_with_time = [(f, get_file_timestamp(f)) for f in files]
        files_with_time = [(f, t) for f, t in files_with_time if t is not None]
        files_with_time.sort(key=lambda x: x[1], reverse=True)
        
        # Xóa file cũ
        files_to_delete = files_with_time[keep_count:]
        deleted_count = 0
        
        for file_path, _ in files_to_delete:
            try:
                Path(file_path).unlink()
                deleted_count += 1
            except Exception:
                pass
        
        return deleted_count
        
    except Exception:
        return 0


def create_timestamped_filename(base_name: str, extension: str, timestamp_format: str = "%Y%m%d_%H%M%S") -> str:
    """
    Tạo tên file với timestamp
    
    Args:
        base_name (str): Tên file cơ bản
        extension (str): Phần mở rộng (với hoặc không có dấu chấm)
        timestamp_format (str): Format timestamp
        
    Returns:
        str: Tên file với timestamp
    """
    timestamp = datetime.now().strftime(timestamp_format)
    if not extension.startswith('.'):
        extension = f'.{extension}'
    
    return f"{base_name}_{timestamp}{extension}"


def validate_file_access(file_path: str, check_read: bool = True, check_write: bool = False) -> bool:
    """
    Kiểm tra quyền truy cập file
    
    Args:
        file_path (str): Đường dẫn file
        check_read (bool): Kiểm tra quyền đọc
        check_write (bool): Kiểm tra quyền ghi
        
    Returns:
        bool: True nếu có đủ quyền
    """
    try:
        path = Path(file_path)
        
        if not path.exists():
            return False
        
        if check_read and not os.access(path, os.R_OK):
            return False
            
        if check_write and not os.access(path, os.W_OK):
            return False
        
        return True
        
    except Exception:
        return False


def get_directory_info(directory: str) -> dict:
    """
    Lấy thông tin thư mục
    
    Args:
        directory (str): Đường dẫn thư mục
        
    Returns:
        dict: Thông tin thư mục
    """
    try:
        path = Path(directory)
        if not path.exists():
            return {'exists': False}
        
        files = [f for f in path.iterdir() if f.is_file()]
        dirs = [d for d in path.iterdir() if d.is_dir()]
        
        total_size = sum(f.stat().st_size for f in files)
        
        return {
            'exists': True,
            'file_count': len(files),
            'dir_count': len(dirs),
            'total_size': total_size,
            'total_size_formatted': format_file_size(total_size),
            'files': [f.name for f in files],
            'directories': [d.name for d in dirs]
        }
        
    except Exception as e:
        return {'exists': False, 'error': str(e)}
