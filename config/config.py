"""
Cấu hình Google Drive và Google Sheets API
Author: Assistant
Date: 2025-07-26
"""

import os
from typing import Dict, List, Optional

# =============================================================================
# GOOGLE API CONFIGURATION
# =============================================================================

# Đường dẫn đến service account JSON file
SERVICE_ACCOUNT_FILE = os.path.join(os.path.dirname(__file__), 'service_account.json')

# Scopes cần thiết cho Google APIs
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

# =============================================================================
# GOOGLE DRIVE CONFIGURATION
# =============================================================================

# ID của Google Drive folder chứa file input
# Có thể lấy từ URL: https://drive.google.com/drive/folders/FOLDER_ID_HERE
DRIVE_INPUT_FOLDER_ID = ""  # TODO: Cập nhật ID folder input

# ID của Google Drive folder để lưu file output
DRIVE_OUTPUT_FOLDER_ID = ""  # TODO: Cập nhật ID folder output

# Tên các file input trên Google Drive
DRIVE_FILES = {
    'students': 'Danh sach hoc sinh.xlsx',
    'teachers': 'DS tài khoản giáo viên.xlsx',
    'template': 'Template_Export.xlsx'
}

# =============================================================================
# GOOGLE SHEETS CONFIGURATION
# =============================================================================

# ID của Google Sheets chứa dữ liệu học sinh
# Có thể lấy từ URL: https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit
STUDENTS_SHEET_ID = ""  # TODO: Cập nhật Sheet ID học sinh

# ID của Google Sheets chứa dữ liệu giáo viên
TEACHERS_SHEET_ID = ""  # TODO: Cập nhật Sheet ID giáo viên

# Tên các sheet trong Google Sheets
SHEET_NAMES = {
    'students': 'Danh sách HS toàn trường',
    'teachers': 'GIAO-VIEN'
}

# =============================================================================
# LOCAL CONFIGURATION
# =============================================================================

# Đường dẫn thư mục local để sync dữ liệu
LOCAL_SYNC_FOLDER = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'sync')

# Thời gian timeout cho API calls (seconds)
API_TIMEOUT = 30

# =============================================================================
# MAPPING CONFIGURATION
# =============================================================================

# Mapping cột từ Google Sheets sang template
STUDENTS_COLUMN_MAPPING = {
    'STT': 'STT',
    'Mã học sinh': 'Mã học sinh',
    'Họ và tên': 'Họ và tên',
    'Ngày sinh': 'Ngày sinh',
    'Khối': 'Khối',
    'Lớp': 'Lớp chính',
    'Tài khoản': 'Tài khoản',
    'Mật khẩu lần đầu': 'Mật khẩu lần đầu',
    'Mã đăng nhập cho PH': 'Mã đăng nhập cho PH',
    'Thời gian tạo': 'Thời gian tạo'
}

TEACHERS_COLUMN_MAPPING = {
    'STT': 'STT',
    'Tên giáo viên': 'Tên giáo viên',
    'Ngày sinh': 'Ngày sinh',
    'Tên đăng nhập': 'Tên đăng nhập',
    'Mật khẩu đăng nhập lần đầu': 'Mật khẩu đăng nhập lần đầu',
    'CCCD/CMND': 'CCCD/CMND',
    'Thời gian tạo': 'Thời gian tạo'
}

# =============================================================================
# VALIDATION FUNCTIONS
# =============================================================================

def validate_config() -> Dict[str, bool]:
    """
    Kiểm tra tính hợp lệ của configuration
    
    Returns:
        Dict[str, bool]: Kết quả validation cho từng thành phần
    """
    results = {
        'service_account_file': os.path.exists(SERVICE_ACCOUNT_FILE),
        'drive_input_folder': bool(DRIVE_INPUT_FOLDER_ID),
        'drive_output_folder': bool(DRIVE_OUTPUT_FOLDER_ID),
        'students_sheet': bool(STUDENTS_SHEET_ID),
        'teachers_sheet': bool(TEACHERS_SHEET_ID),
        'local_sync_folder': True  # Sẽ được tạo tự động
    }
    
    return results

def print_config_status():
    """In trạng thái cấu hình hiện tại"""
    print("\n📋 TRẠNG THÁI CẤU HÌNH GOOGLE API:")
    print("=" * 50)
    
    validation = validate_config()
    
    for key, status in validation.items():
        status_icon = "✅" if status else "❌"
        key_display = key.replace('_', ' ').title()
        print(f"{status_icon} {key_display}: {'OK' if status else 'Chưa cấu hình'}")
    
    if not validation['service_account_file']:
        print(f"\n⚠️  File service account không tồn tại: {SERVICE_ACCOUNT_FILE}")
        print("   Hãy tải file JSON từ Google Cloud Console và đặt vào thư mục config/")
    
    if not validation['drive_input_folder']:
        print("\n⚠️  Chưa cấu hình DRIVE_INPUT_FOLDER_ID")
        print("   Hãy cập nhật ID folder input trong config.py")
    
    if not validation['drive_output_folder']:
        print("\n⚠️  Chưa cấu hình DRIVE_OUTPUT_FOLDER_ID")
        print("   Hãy cập nhật ID folder output trong config.py")
    
    if not validation['students_sheet']:
        print("\n⚠️  Chưa cấu hình STUDENTS_SHEET_ID")
        print("   Hãy cập nhật ID Google Sheets học sinh trong config.py")
    
    if not validation['teachers_sheet']:
        print("\n⚠️  Chưa cấu hình TEACHERS_SHEET_ID")
        print("   Hãy cập nhật ID Google Sheets giáo viên trong config.py")
    
    print("\n📝 Hướng dẫn chi tiết xem file README.md trong thư mục config/")

def get_sample_config():
    """Trả về cấu hình mẫu"""
    return {
        'DRIVE_INPUT_FOLDER_ID': '1ABC123def456GHI789jkl',
        'DRIVE_OUTPUT_FOLDER_ID': '1XYZ789abc123DEF456ghi',
        'STUDENTS_SHEET_ID': '1MNO456pqr789STU123vwx',
        'TEACHERS_SHEET_ID': '1QRS789tuv123WXY456zab'
    }

if __name__ == "__main__":
    print_config_status()
    
    print("\n📋 CẤU HÌNH MẪU:")
    print("=" * 50)
    sample = get_sample_config()
    for key, value in sample.items():
        print(f"{key} = '{value}'")
