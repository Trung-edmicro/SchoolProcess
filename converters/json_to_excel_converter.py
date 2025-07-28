"""
JSON to Excel Template Converter (Chuẩn Mode 1)
Chuyển đổi dữ liệu JSON từ OnLuyen API thành file Excel format đúng template Mode 1
Author: Assistant  
Date: 2025-07-26
"""

import json
import pandas as pd
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from datetime import datetime
import os
import shutil
from pathlib import Path


class JSONToExcelTemplateConverter:
    """Converter chuyển JSON sang Excel format theo template chuẩn Mode 1"""
    
    def __init__(self, json_file_path: str, template_path: str = None):
        """
        Khởi tạo converter
        
        Args:
            json_file_path (str): Đường dẫn file JSON
            template_path (str): Đường dẫn template Excel
        """
        self.json_file_path = json_file_path
        self.template_path = template_path or "data/temp/Template_Export.xlsx"
        self.json_data = None
        self.school_name = ""
        self.admin_email = ""
        self.admin_password = ""  # Thêm biến lưu admin password
        self.teachers_df = None
        self.students_df = None
        
        # Thiết lập style cho Excel (chuẩn như base_processor)
        self.thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'), 
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )
        
        # Alignment styles
        self.center_alignment = Alignment(horizontal='center', vertical='center')
        self.left_alignment = Alignment(horizontal='left', vertical='center')
        
        # Font styles
        self.header_font = Font(bold=True, size=11, name='Calibri')
        self.data_font = Font(size=11, name='Calibri')
        
        # Fill styles
        self.header_fill = PatternFill(start_color="D9EDF7", end_color="D9EDF7", fill_type="solid")
        
    def load_json_data(self):
        """Load dữ liệu từ file JSON"""
        try:
            with open(self.json_file_path, 'r', encoding='utf-8') as f:
                self.json_data = json.load(f)
            
            print(f"✅ Đã load JSON data từ: {self.json_file_path}")
            
            # Extract school info từ cấu trúc mới (có thể có metadata hoặc school_info trực tiếp)
            if 'metadata' in self.json_data:
                # Cấu trúc cũ với metadata
                metadata = self.json_data.get('metadata', {})
                school_info = metadata.get('school_info', {})
                self.admin_password = metadata.get('admin_password', '123456')
            else:
                # Cấu trúc mới với school_info trực tiếp
                school_info = self.json_data.get('school_info', {})
                # Sửa: password được lưu trong school_info.password, không phải admin_password
                self.admin_password = school_info.get('password', '123456')
            
            self.school_name = school_info.get('name', 'Unknown School')
            self.admin_email = school_info.get('admin', '')
            
            print(f"📋 Tên trường: {self.school_name}")
            print(f"📧 Admin email: {self.admin_email}")
            print(f"🔑 Admin password: {'*' * len(self.admin_password) if self.admin_password else 'N/A'}")
            
            return True
            
        except Exception as e:
            print(f"❌ Lỗi khi load JSON: {e}")
            return False
    
    def extract_teachers_data(self):
        """Trích xuất dữ liệu giáo viên từ JSON - hỗ trợ cả workflow và filtered format"""
        try:
            print("   📊 Đang trích xuất dữ liệu giáo viên...")
            
            # Xử lý các format khác nhau của dữ liệu teachers
            teachers_raw = self.json_data.get('teachers', [])
            
            if isinstance(teachers_raw, dict):
                # Format workflow thường: {'data': [...]}
                teachers_data = teachers_raw.get('data', [])
                print(f"   📋 Phát hiện format workflow: {len(teachers_data)} giáo viên")
            elif isinstance(teachers_raw, list):
                # Format filtered: [...] trực tiếp
                teachers_data = teachers_raw
                print(f"   📋 Phát hiện format filtered: {len(teachers_data)} giáo viên")
            else:
                print(f"   ❌ Format dữ liệu teachers không được hỗ trợ: {type(teachers_raw)}")
                return False
            
            if not teachers_data:
                print("   ⚠️ Không có dữ liệu giáo viên")
                return False
            
            teachers_list = []
            
            for idx, teacher_record in enumerate(teachers_data, 1):
                # Xử lý cấu trúc teacher record linh hoạt
                if isinstance(teacher_record, dict):
                    # Tìm teacherInfo trong các vị trí có thể
                    teacher_info_data = (
                        teacher_record.get('teacherInfo', {}) or 
                        teacher_record.get('userInfo', {}) or 
                        teacher_record
                    )
                    
                    teacher_info = {
                        'STT': idx,
                        'Tên giáo viên': teacher_info_data.get('displayName', ''),
                        'Ngày sinh': teacher_info_data.get('userBirthday', ''),
                        'Tên đăng nhập': teacher_info_data.get('userName', ''),
                        'Mật khẩu đăng nhập lần đầu': teacher_info_data.get('pwd', '')
                    }
                    teachers_list.append(teacher_info)
            
            self.teachers_df = pd.DataFrame(teachers_list)
            print(f"✅ Đã trích xuất {len(teachers_list)} giáo viên")
            
            return True
            
        except Exception as e:
            print(f"❌ Lỗi khi trích xuất dữ liệu giáo viên: {e}")
            return False
    
    def extract_students_data(self):
        """Trích xuất dữ liệu học sinh từ JSON - hỗ trợ cả workflow và filtered format"""
        try:
            print("   📊 Đang trích xuất dữ liệu học sinh...")
            
            # Xử lý các format khác nhau của dữ liệu students
            students_raw = self.json_data.get('students', [])
            
            if isinstance(students_raw, dict):
                # Format workflow thường: {'data': [...]}
                students_data = students_raw.get('data', [])
                print(f"   📋 Phát hiện format workflow: {len(students_data)} học sinh")
            elif isinstance(students_raw, list):
                # Format filtered: [...] trực tiếp
                students_data = students_raw
                print(f"   📋 Phát hiện format filtered: {len(students_data)} học sinh")
            else:
                print(f"   ❌ Format dữ liệu students không được hỗ trợ: {type(students_raw)}")
                return False
            
            if not students_data:
                print("   ⚠️ Không có dữ liệu học sinh")
                return False
            
            students_list = []
            
            for idx, student_record in enumerate(students_data, 1):
                # Xử lý cấu trúc student record linh hoạt
                if isinstance(student_record, dict):
                    # Tìm userInfo trong các vị trí có thể
                    user_info = (
                        student_record.get('userInfo', {}) or 
                        student_record.get('studentInfo', {}) or 
                        student_record
                    )
                    
                    student_info = {
                        'STT': idx,
                        'Họ và tên': user_info.get('displayName', ''),
                        'Ngày sinh': user_info.get('userBirthday', ''),
                        'Lớp chính': student_record.get('grade', ''),
                        'Tài khoản': user_info.get('userName', ''),
                        'Mật khẩu lần đầu': user_info.get('pwd', ''),
                        'Mã đăng nhập cho PH': user_info.get('codePin', '')
                    }
                    students_list.append(student_info)
            
            self.students_df = pd.DataFrame(students_list)
            print(f"✅ Đã trích xuất {len(students_list)} học sinh")
            
            return True
            
        except Exception as e:
            print(f"❌ Lỗi khi trích xuất dữ liệu học sinh: {e}")
            return False
    
    def apply_border_to_sheet(self, sheet, max_row: int, max_col: int, center_columns: list = None):
        """
        Áp dụng border và căn giữa cho các ô có dữ liệu trong sheet (chuẩn như base_processor)
        
        Args:
            sheet: Sheet cần áp dụng border
            max_row (int): Số hàng tối đa có dữ liệu
            max_col (int): Số cột tối đa có dữ liệu
            center_columns (list): Danh sách các cột cần căn giữa (1-based)
        """
        try:
            if center_columns is None:
                center_columns = []
            
            # Áp dụng border cho tất cả các ô có dữ liệu
            for row in range(1, max_row + 1):
                for col in range(1, max_col + 1):
                    cell = sheet.cell(row=row, column=col)
                    # Chỉ áp dụng border cho ô có dữ liệu hoặc là header
                    if cell.value is not None or row == 1:
                        cell.border = self.thin_border
                        
                        # Căn giữa cho các cột được chỉ định
                        if col in center_columns:
                            cell.alignment = self.center_alignment
                        else:
                            cell.alignment = self.left_alignment
            
            print(f"✅ Đã áp dụng border và căn giữa cho sheet {sheet.title}")
            
        except Exception as e:
            print(f"❌ Lỗi khi áp dụng style cho sheet {sheet.title}: {e}")

    def copy_template(self, output_path: str):
        """Copy template làm base cho file output"""
        try:
            shutil.copy2(self.template_path, output_path)
            print(f"✅ Đã copy template thành: {output_path}")
            return True
        except Exception as e:
            print(f"❌ Lỗi khi copy template: {e}")
            return False
    
    def update_admin_sheet(self, workbook):
        """Cập nhật sheet ADMIN với thông tin trường, admin và HT/HP - HOÀN TOÀN KHÔNG THAY ĐỔI cấu trúc template"""
        try:
            admin_sheet = workbook['ADMIN']
            
            # Xử lý merged cells để tránh lỗi read-only
            if admin_sheet.merged_cells:
                merged_ranges = list(admin_sheet.merged_cells.ranges)
                for merged_range in merged_ranges:
                    try:
                        admin_sheet.unmerge_cells(str(merged_range))
                    except:
                        pass  # Bỏ qua nếu không thể unmerge
            
            print(f"🔍 DEBUG: Analyzing template structure...")
            print(f"   A1 value: '{admin_sheet['A1'].value}'")
            print(f"   A2 value: '{admin_sheet['A2'].value}'") 
            print(f"   A3 value: '{admin_sheet['A3'].value}'")
            print(f"   A4 value: '{admin_sheet['A4'].value}'")
            print(f"   C1 value: '{admin_sheet['C1'].value}'")
            print(f"   D1 value: '{admin_sheet['D1'].value}'")
            
            # Điền dữ liệu vào các dòng có sẵn trong template (KHÔNG THAY ĐỔI cấu trúc)
            accounts_updated = 0
            
            # Row 2 (A2): Admin - điền vào cột C2 và D2
            try:
                if (admin_sheet['A2'].value and 
                    str(admin_sheet['A2'].value).strip().upper() == 'ADMIN'):
                    admin_sheet['C2'] = self.admin_email  # Cột C: Tài khoản
                    admin_sheet['D2'] = self.admin_password  # Cột D: Mật khẩu
                    accounts_updated += 1
                    print(f"   ✅ Đã điền dữ liệu Admin vào row 2 (A2='Admin')")
                    print(f"      C2: {self.admin_email}")
                    print(f"      D2: {self.admin_password}")
                else:
                    print(f"   ⚠️ Row 2 không có 'Admin': A2='{admin_sheet['A2'].value}'")
            except Exception as e:
                print(f"   ⚠️ Không thể điền dữ liệu Admin: {e}")
            
            # Xử lý HT/HP từ JSON data
            ht_hp_info = self.json_data.get('ht_hp_info', {})
            ht_list = ht_hp_info.get('ht', [])
            hp_list = ht_hp_info.get('hp', [])
            
            print(f"   🔍 HT/HP data: {len(ht_list)} HT, {len(hp_list)} HP")
            
            # Row 3 (A3): Hiệu Trưởng - điền vào cột C3 và D3
            try:
                a3_value = str(admin_sheet['A3'].value).strip() if admin_sheet['A3'].value else ""
                if a3_value.upper() in ['HIỆU TRƯỞNG', 'HIEU TRUONG']:
                    if ht_list:
                        ht = ht_list[0]  # CHỈ lấy HT đầu tiên
                        admin_sheet['C3'] = ht.get('userName', '')
                        admin_sheet['D3'] = ht.get('pwd', '')
                        accounts_updated += 1
                        print(f"   ✅ Đã điền dữ liệu Hiệu Trưởng vào row 3 (A3='{a3_value}')")
                        print(f"      C3: {ht.get('userName', '')}")
                        print(f"      D3: {ht.get('pwd', '')}")
                        
                        if len(ht_list) > 1:
                            print(f"   ⚠️ Có {len(ht_list)} Hiệu trường, template chỉ hỗ trợ 1")
                    else:
                        print(f"   📋 Row 3 có '{a3_value}' nhưng không có dữ liệu HT")
                else:
                    print(f"   ⚠️ Row 3 không có 'Hiệu Trưởng': A3='{a3_value}'")
            except Exception as e:
                print(f"   ⚠️ Không thể điền dữ liệu Hiệu Trưởng: {e}")
            
            # Row 4 (A4): Hiệu Phó - điền vào cột C4 và D4
            try:
                a4_value = str(admin_sheet['A4'].value).strip() if admin_sheet['A4'].value else ""
                if a4_value.upper() in ['HIỆU PHÓ', 'HIEU PHO']:
                    if hp_list:
                        hp = hp_list[0]  # CHỈ lấy HP đầu tiên
                        admin_sheet['C4'] = hp.get('userName', '')
                        admin_sheet['D4'] = hp.get('pwd', '')
                        accounts_updated += 1
                        print(f"   ✅ Đã điền dữ liệu Hiệu Phó vào row 4 (A4='{a4_value}')")
                        print(f"      C4: {hp.get('userName', '')}")
                        print(f"      D4: {hp.get('pwd', '')}")
                        
                        if len(hp_list) > 1:
                            print(f"   ⚠️ Có {len(hp_list)} Hiệu phó, template chỉ hỗ trợ 1")
                    else:
                        print(f"   📋 Row 4 có '{a4_value}' nhưng không có dữ liệu HP")
                else:
                    print(f"   ⚠️ Row 4 không có 'Hiệu Phó': A4='{a4_value}'")
            except Exception as e:
                print(f"   ⚠️ Không thể điền dữ liệu Hiệu Phó: {e}")
            
            # HOÀN TOÀN KHÔNG THÊM DÒNG MỚI - Template structure được giữ nguyên
            
            # Format các ô đã điền dữ liệu (CHỈ format data cells)
            for row in range(2, 5):  # Row 2, 3, 4 (data rows)
                for col in ['C', 'D']:  # Chỉ format cột C và D
                    try:
                        cell = admin_sheet[f'{col}{row}']
                        if hasattr(cell, 'font') and cell.value:
                            cell.font = self.data_font
                            if col == 'D':  # Center align cho mật khẩu
                                cell.alignment = self.center_alignment
                            else:
                                cell.alignment = self.left_alignment
                    except:
                        pass  # Bỏ qua nếu không thể format

            # Điều chỉnh column widths một cách an toàn
            try:
                admin_sheet.column_dimensions['A'].width = 15  # Vai trò  
                admin_sheet.column_dimensions['C'].width = 35  # Tên đăng nhập
                admin_sheet.column_dimensions['D'].width = 25  # Mật khẩu
            except:
                pass  # Bỏ qua nếu không thể điều chỉnh width
            
            print(f"✅ Đã cập nhật sheet ADMIN với {accounts_updated} tài khoản")
            print(f"   📋 Format: Row 2=Admin, Row 3=HT, Row 4=HP")
            print(f"   🏗️ Cột C=Tài khoản, Cột D=Mật khẩu lần đầu")
            print(f"   ✅ HOÀN TOÀN GIỮ NGUYÊN template structure")
            return True
            
        except Exception as e:
            print(f"❌ Lỗi khi cập nhật sheet ADMIN: {e}")
            return False
    
    def fill_teachers_sheet(self, workbook):
        """Điền dữ liệu giáo viên vào sheet GIAO-VIEN"""
        try:
            if self.teachers_df is None or self.teachers_df.empty:
                print("⚠️ Không có dữ liệu giáo viên để điền")
                return True
                
            teachers_sheet = workbook['GIAO-VIEN']
            
            # Clear existing data (keep headers)
            teachers_sheet.delete_rows(2, teachers_sheet.max_row)
            
            # Điền dữ liệu từ row 2
            for idx, row in self.teachers_df.iterrows():
                row_num = idx + 2
                
                teachers_sheet[f'A{row_num}'] = row['STT']
                teachers_sheet[f'B{row_num}'] = row['Tên giáo viên']
                teachers_sheet[f'C{row_num}'] = row['Ngày sinh']
                teachers_sheet[f'D{row_num}'] = row['Tên đăng nhập']
                teachers_sheet[f'E{row_num}'] = row['Mật khẩu đăng nhập lần đầu']
            
            # Áp dụng border và alignment chuẩn như base_processor
            max_data_row = len(self.teachers_df) + 1  # +1 cho header
            max_data_col = 5  # 5 cột: STT, Tên, Ngày sinh, Tên đăng nhập, Mật khẩu
            
            # Các cột cần căn giữa: STT (1), Ngày sinh (3), Mật khẩu (5)
            center_columns = [1, 3, 5]
            self.apply_border_to_sheet(teachers_sheet, max_data_row, max_data_col, center_columns)
            
            # Auto-adjust column widths theo chuẩn Mode 1
            column_widths = {
                'A': 8,   # STT
                'B': 30,  # Tên giáo viên - rộng hơn cho tên dài
                'C': 15,  # Ngày sinh
                'D': 40,  # Tên đăng nhập - rộng hơn cho email
                'E': 25   # Mật khẩu
            }
            for col, width in column_widths.items():
                teachers_sheet.column_dimensions[col].width = width
                
            # Set row height để text hiển thị đẹp
            for row_num in range(1, len(self.teachers_df) + 2):
                teachers_sheet.row_dimensions[row_num].height = 20
            
            print(f"✅ Đã điền {len(self.teachers_df)} giáo viên vào sheet GIAO-VIEN")
            return True
            
        except Exception as e:
            print(f"❌ Lỗi khi điền dữ liệu giáo viên: {e}")
            return False
    
    def fill_students_sheet(self, workbook):
        """Điền dữ liệu học sinh vào sheet HOC-SINH"""
        try:
            if self.students_df is None or self.students_df.empty:
                print("⚠️ Không có dữ liệu học sinh để điền")
                return True
                
            students_sheet = workbook['HOC-SINH']
            
            # Clear existing data (keep headers)
            students_sheet.delete_rows(2, students_sheet.max_row)
            
            # Điền dữ liệu từ row 2
            for idx, row in self.students_df.iterrows():
                row_num = idx + 2
                
                students_sheet[f'A{row_num}'] = row['STT']
                students_sheet[f'B{row_num}'] = row['Họ và tên']
                students_sheet[f'C{row_num}'] = row['Ngày sinh']
                students_sheet[f'D{row_num}'] = row['Lớp chính']
                students_sheet[f'E{row_num}'] = row['Tài khoản']
                students_sheet[f'F{row_num}'] = row['Mật khẩu lần đầu']
                students_sheet[f'G{row_num}'] = row['Mã đăng nhập cho PH']
            
            # Áp dụng border và alignment chuẩn như base_processor
            max_data_row = len(self.students_df) + 1  # +1 cho header
            max_data_col = 7  # 7 cột: STT, Họ tên, Ngày sinh, Lớp, Tài khoản, Mật khẩu, Mã PH
            
            # Các cột cần căn giữa: STT (1), Ngày sinh (3), Lớp (4), Mật khẩu (6), Mã PH (7)
            center_columns = [1, 3, 4, 6, 7]
            self.apply_border_to_sheet(students_sheet, max_data_row, max_data_col, center_columns)
            
            # Auto-adjust column widths theo chuẩn Mode 1
            column_widths = {
                'A': 8,   # STT
                'B': 30,  # Họ và tên - rộng hơn cho tên dài
                'C': 15,  # Ngày sinh
                'D': 20,  # Lớp chính
                'E': 40,  # Tài khoản - rộng hơn cho username/email
                'F': 20,  # Mật khẩu lần đầu
                'G': 20   # Mã đăng nhập cho PH
            }
            for col, width in column_widths.items():
                students_sheet.column_dimensions[col].width = width
                
            # Set row height để text hiển thị đẹp
            for row_num in range(1, len(self.students_df) + 2):
                students_sheet.row_dimensions[row_num].height = 20
            
            print(f"✅ Đã điền {len(self.students_df)} học sinh vào sheet HOC-SINH")
            return True
            
        except Exception as e:
            print(f"❌ Lỗi khi điền dữ liệu học sinh: {e}")
            return False
    
    def create_excel_output(self, output_path: str = None):
        """Tạo file Excel output theo template chuẩn Mode 1"""
        try:
            if output_path is None:
                # Tạo tên file với format "Export_Tên trường.xlsx"
                safe_school_name = "".join(c for c in self.school_name if c.isalnum() or c in (' ', '-')).strip()
                output_path = f"data/output/Export_{safe_school_name}.xlsx"
            
            # Đảm bảo thư mục output tồn tại
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Copy template làm base
            if not self.copy_template(output_path):
                return None
            
            # Load workbook
            workbook = load_workbook(output_path)
            
            # Cập nhật các sheet
            if not self.update_admin_sheet(workbook):
                return None
                
            if not self.fill_teachers_sheet(workbook):
                return None
                
            if not self.fill_students_sheet(workbook):
                return None
            
            # Save workbook
            workbook.save(output_path)
            workbook.close()
            
            print(f"🎉 Đã tạo thành công file Excel: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"❌ Lỗi khi tạo file Excel: {e}")
            return None
    
    def convert(self, output_path: str = None):
        """Thực hiện chuyển đổi toàn bộ"""
        print("🚀 BẮT ĐẦU CHUYỂN ĐỔI JSON SANG EXCEL (TEMPLATE MODE 1)")
        print("=" * 60)
        
        # Load JSON data
        if not self.load_json_data():
            return None
        
        # Extract data
        if not self.extract_teachers_data():
            return None
            
        if not self.extract_students_data():
            return None
        
        # Create Excel output
        result_path = self.create_excel_output(output_path)
        
        if result_path:
            print("\n" + "=" * 60)
            print("🎊 CHUYỂN ĐỔI HOÀN TẤT!")
            print(f"📄 File Excel đầu ra: {result_path}")
            print(f"👥 Số giáo viên: {len(self.teachers_df) if self.teachers_df is not None else 0}")
            print(f"🎓 Số học sinh: {len(self.students_df) if self.students_df is not None else 0}")
            print(f"🏫 Trường: {self.school_name}")
        
        return result_path


def convert_json_to_excel_template(json_file_path: str, output_path: str = None, template_path: str = None):
    """
    Hàm tiện ích để chuyển đổi JSON sang Excel template format
    
    Args:
        json_file_path (str): Đường dẫn file JSON
        output_path (str): Đường dẫn file Excel output (optional)
        template_path (str): Đường dẫn template Excel (optional)
    
    Returns:
        str: Đường dẫn file Excel output nếu thành công, None nếu lỗi
    """
    converter = JSONToExcelTemplateConverter(json_file_path, template_path)
    return converter.convert(output_path)


if __name__ == "__main__":
    # Test với file JSON workflow
    print("🧪 TEST JSON TO EXCEL TEMPLATE CONVERTER")
    print("=" * 50)
    
    # Đường dẫn file JSON (sử dụng file mới nhất)
    json_files = [
        "data/output/workflow_data_GDNN - GDTX TP Chí Linh_20250726_160141.json"
    ]
    
    for json_file in json_files:
        if os.path.exists(json_file):
            print(f"\n🔄 Đang chuyển đổi: {json_file}")
            result = convert_json_to_excel_template(json_file)
            
            if result:
                print(f"✅ Thành công: {result}")
            else:
                print("❌ Chuyển đổi thất bại")
        else:
            print(f"❌ File không tồn tại: {json_file}")
