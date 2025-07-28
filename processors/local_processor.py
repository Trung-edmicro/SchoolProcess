"""
Local Data Processor
Processor xử lý dữ liệu từ file Excel local
Author: Assistant
Date: 2025-07-26
"""

import pandas as pd
import os
from .base_processor import BaseDataProcessor


class LocalDataProcessor(BaseDataProcessor):
    """Processor xử lý dữ liệu từ file Excel local"""
    
    def __init__(self, input_folder: str, temp_folder: str, output_folder: str = None):
        """
        Khởi tạo LocalDataProcessor
        
        Args:
            input_folder (str): Thư mục chứa file input
            temp_folder (str): Thư mục chứa template
            output_folder (str): Thư mục output (mặc định là temp_folder)
        """
        super().__init__(input_folder, temp_folder, output_folder)
    
    def get_processor_name(self) -> str:
        """Trả về tên processor"""
        return "LOCAL FILE PROCESSOR"
    
    def load_students_data(self) -> pd.DataFrame:
        """
        Load dữ liệu học sinh từ file Excel local
        
        Returns:
            pd.DataFrame: Dữ liệu học sinh
        """
        try:
            # Kiểm tra file tồn tại
            if not os.path.exists(self.student_file):
                raise FileNotFoundError(f"File học sinh không tồn tại: {self.student_file}")
            
            # Đọc dữ liệu từ hàng 6 (vì có multi-header)
            print(f"📖 Đang đọc file học sinh: {self.student_file}")
            
            # Đọc với header tại hàng 6 (index 5)
            df = pd.read_excel(
                self.student_file, 
                sheet_name="Danh sách HS toàn trường",
                header=5,  # Header tại hàng 6
                engine='openpyxl'
            )
            
            # Loại bỏ các hàng trống
            df = df.dropna(how='all')
            
            # Đổi tên các cột để đồng nhất
            column_mapping = {
                'STT': 'STT',
                'Họ và tên': 'Họ và tên',
                'Ngày sinh': 'Ngày sinh',
                'Lớp chính': 'Lớp chính',
                'Tài khoản': 'Tài khoản',
                'Mật khẩu lần đầu': 'Mật khẩu lần đầu',
                'Mã đăng nhập cho PH': 'Mã đăng nhập cho PH'
            }
            
            # Rename columns nếu tồn tại
            for old_col, new_col in column_mapping.items():
                if old_col in df.columns:
                    df = df.rename(columns={old_col: new_col})
            
            # Đảm bảo có đủ các cột cần thiết
            required_columns = ['STT', 'Họ và tên', 'Ngày sinh', 'Lớp chính', 'Tài khoản', 'Mật khẩu lần đầu', 'Mã đăng nhập cho PH']
            for col in required_columns:
                if col not in df.columns:
                    df[col] = ''  # Tạo cột trống nếu không có
            
            # Lọc chỉ các cột cần thiết
            df = df[required_columns]
            
            # Loại bỏ các hàng có STT không phải là số
            df = df[pd.to_numeric(df['STT'], errors='coerce').notna()]
            
            # Reset index
            df = df.reset_index(drop=True)
            
            self.students_data = df
            print(f"✅ Đã load {len(df)} học sinh từ file local")
            
            return df
            
        except Exception as e:
            print(f"❌ Lỗi khi đọc file học sinh: {e}")
            self.students_data = pd.DataFrame()
            return pd.DataFrame()
    
    def load_teachers_data(self) -> pd.DataFrame:
        """
        Load dữ liệu giáo viên từ file Excel local
        
        Returns:
            pd.DataFrame: Dữ liệu giáo viên
        """
        try:
            # Kiểm tra file tồn tại
            if not os.path.exists(self.teacher_file):
                raise FileNotFoundError(f"File giáo viên không tồn tại: {self.teacher_file}")
            
            # Đọc dữ liệu giáo viên
            print(f"📖 Đang đọc file giáo viên: {self.teacher_file}")
            
            # Thử đọc với header mặc định trước
            df = pd.read_excel(
                self.teacher_file,
                sheet_name=0,  # Sheet đầu tiên
                engine='openpyxl'
            )
            
            # Loại bỏ các hàng trống
            df = df.dropna(how='all')
            
            # Đổi tên các cột để đồng nhất
            column_mapping = {
                'STT': 'STT',
                'Tên giáo viên': 'Tên giáo viên',
                'Ngày sinh': 'Ngày sinh', 
                'Tên đăng nhập': 'Tên đăng nhập',
                'Mật khẩu đăng nhập lần đầu': 'Mật khẩu đăng nhập lần đầu'
            }
            
            # Rename columns nếu tồn tại
            for old_col, new_col in column_mapping.items():
                if old_col in df.columns:
                    df = df.rename(columns={old_col: new_col})
            
            # Đảm bảo có đủ các cột cần thiết
            required_columns = ['STT', 'Tên giáo viên', 'Ngày sinh', 'Tên đăng nhập', 'Mật khẩu đăng nhập lần đầu']
            for col in required_columns:
                if col not in df.columns:
                    df[col] = ''  # Tạo cột trống nếu không có
            
            # Lọc chỉ các cột cần thiết
            df = df[required_columns]
            
            # Loại bỏ các hàng có STT không phải là số
            df = df[pd.to_numeric(df['STT'], errors='coerce').notna()]
            
            # Reset index
            df = df.reset_index(drop=True)
            
            self.teachers_data = df
            print(f"✅ Đã load {len(df)} giáo viên từ file local")
            
            return df
            
        except Exception as e:
            print(f"❌ Lỗi khi đọc file giáo viên: {e}")
            self.teachers_data = pd.DataFrame()
            return pd.DataFrame()
    
    def process_local_files(self) -> str:
        """
        Xử lý toàn bộ quy trình với file local
        
        Returns:
            str: Đường dẫn file output
        """
        print("\n🏠 CHẾ ĐỘ 1: XỬ LÝ VỚI FILE LOCAL")
        print("=" * 50)
        
        # Kiểm tra các file input tồn tại
        missing_files = []
        if not os.path.exists(self.student_file):
            missing_files.append(f"File học sinh: {self.student_file}")
        if not os.path.exists(self.teacher_file):
            missing_files.append(f"File giáo viên: {self.teacher_file}")
        if not os.path.exists(self.template_file):
            missing_files.append(f"File template: {self.template_file}")
        
        if missing_files:
            print("❌ Các file sau không tồn tại:")
            for file in missing_files:
                print(f"   - {file}")
            return ""
        
        # Xử lý mapping dữ liệu
        return self.process_all()
    
    def validate_input_files(self) -> bool:
        """
        Kiểm tra tính hợp lệ của các file input
        
        Returns:
            bool: True nếu tất cả file hợp lệ
        """
        try:
            print("\n🔍 KIỂM TRA FILE INPUT:")
            
            # Kiểm tra file học sinh
            if os.path.exists(self.student_file):
                df_students = pd.read_excel(self.student_file, sheet_name="Danh sách HS toàn trường", header=5)
                students_count = len(df_students.dropna(how='all'))
                print(f"   📚 File học sinh: ✅ ({students_count} học sinh)")
            else:
                print(f"   📚 File học sinh: ❌ Không tồn tại")
                return False
            
            # Kiểm tra file giáo viên
            if os.path.exists(self.teacher_file):
                df_teachers = pd.read_excel(self.teacher_file, sheet_name=0)
                teachers_count = len(df_teachers.dropna(how='all'))
                print(f"   👨‍🏫 File giáo viên: ✅ ({teachers_count} giáo viên)")
            else:
                print(f"   👨‍🏫 File giáo viên: ❌ Không tồn tại")
                return False
            
            # Kiểm tra file template
            if os.path.exists(self.template_file):
                print(f"   📋 File template: ✅")
            else:
                print(f"   📋 File template: ❌ Không tồn tại")
                return False
            
            print("✅ Tất cả file input đều hợp lệ")
            return True
            
        except Exception as e:
            print(f"❌ Lỗi khi kiểm tra file input: {e}")
            return False
