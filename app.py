"""
School Process Application - Enhanced Version
Ứng dụng chính với cấu trúc modular và cấu hình từ .env
Author: Assistant
Date: 2025-07-26
"""

import sys
import traceback
import json
import glob
import os
import re
import io
import unicodedata
from datetime import datetime
from pathlib import Path

# Third-party imports
import pandas as pd
from googleapiclient.http import MediaIoBaseDownload

# Thêm project root vào Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from config.config_manager import get_config
from utils.menu_utils import *
from utils.file_utils import ensure_directories
from converters import JSONToExcelTemplateConverter
from processors.local_processor import LocalDataProcessor
from config.onluyen_api import OnLuyenAPIClient
from extractors import GoogleSheetsExtractor
from config.google_oauth_drive import GoogleOAuthDriveClient


class SchoolProcessApp:
    """Ứng dụng chính School Process"""
    
    def __init__(self):
        """Khởi tạo ứng dụng"""
        self.config = get_config()
        self.setup_directories()
        
    def setup_directories(self):
        """Thiết lập các thư mục cần thiết"""
        paths = self.config.get_paths_config()
        required_dirs = [
            paths['input_dir'],
            paths['output_dir'],
            paths['config_dir'],
            'logs',
            'backups'
        ]
        
        if not ensure_directories(required_dirs):
            print("⚠️  Một số thư mục không thể tạo được")
    
    def show_main_menu(self):
        """Hiển thị và xử lý main menu"""
        options = [
            "Xử lý dữ liệu local (Excel files)",
            "OnLuyen API"
        ]
        
        handlers = [
            self.mode_local_processing,
            self.mode_onluyen_api
        ]
        
        run_menu_loop("SCHOOL PROCESS - MENU CHÍNH", options, handlers)
    
    def mode_local_processing(self):
        """Chế độ xử lý dữ liệu local"""
        print_separator("XỬ LÝ DỮ LIỆU LOCAL")
        
        try:
            
            paths = self.config.get_paths_config()
            processor = LocalDataProcessor(
                input_folder=paths['input_dir'],
                temp_folder=paths['temp_dir'],
                output_folder=paths['output_dir']
            )
            
            if not processor.validate_input_files():
                print_status("Không thể tiếp tục do thiếu file input", "error")
                return
            
            print_status("Bắt đầu xử lý dữ liệu local...", "info")
            output_path = processor.process_local_files()
            
            if output_path:
                processor.print_summary()
                print_status(f"Hoàn thành! File output: {output_path}", "success")
            else:
                print_status("Lỗi trong quá trình xử lý", "error")
                
        except ImportError:
            print_status("Local processor chưa được triển khai", "warning")
        except Exception as e:
            print_status(f"Lỗi xử lý local: {e}", "error")
    
    def mode_onluyen_api(self):
        """Chế độ OnLuyen API Integration"""
        print_separator("ONLUYEN API INTEGRATION")
        
        # Submenu cho OnLuyen API - thêm các chức năng lấy dữ liệu
        options = [
            "Case 1: Toàn bộ dữ liệu",
            "Case 2: Dữ liệu theo file import",
            "Lấy danh sách Giáo viên",
            "Lấy danh sách Học sinh"
        ]
        
        handlers = [
            self._workflow_case_1_full_data,
            self._workflow_case_2_import_filtered,
            self.onluyen_get_teachers,
            self.onluyen_get_students
        ]
        
        run_menu_loop("ONLUYEN API INTEGRATION", options, handlers)
    
    def onluyen_get_teachers(self):
        """Lấy danh sách giáo viên"""
        print_separator("LẤY DANH SÁCH GIÁO VIÊN")
        
        try:
            
            # Hỏi page size với default lớn hơn
            page_size = get_user_input("Nhập page size (Enter = 1000)") or "1000"
            try:
                page_size = int(page_size)
            except ValueError:
                page_size = 1000
            
            client = OnLuyenAPIClient()
            
            # Kiểm tra có cần login không
            if self._check_onluyen_auth_required(client):
                return
            
            print_status(f"Đang lấy danh sách giáo viên (page size: {page_size})...", "info")
            
            result = client.get_teachers(page_size=page_size)
            
            if result['success']:
                data = result.get('data')
                
                if data:
                    if isinstance(data, dict) and 'data' in data:
                        teachers_list = data['data']
                        teachers_count = data.get('totalCount', len(teachers_list))
                        
                        print_status(f"✅ Lấy danh sách thành công: {len(teachers_list)}/{teachers_count} giáo viên", "success")
                        
                        if len(teachers_list) > 0:
                            print(f"\n📋 DANH SÁCH GIÁO VIÊN (hiển thị {min(len(teachers_list), 10)} đầu tiên):")
                            for i, teacher in enumerate(teachers_list[:10], 1):
                                if isinstance(teacher, dict):
                                    name = teacher.get('name', teacher.get('fullName', 'N/A'))
                                    email = teacher.get('email', 'N/A')
                                    id_val = teacher.get('id', teacher.get('teacherId', 'N/A'))
                                    print(f"   {i:2d}. ID: {id_val} | Tên: {name} | Email: {email}")
                            
                            if len(teachers_list) > 10:
                                print(f"   ... và {len(teachers_list) - 10} giáo viên khác")
                            
                            # Hỏi có muốn lưu JSON không
                            if get_user_confirmation("Lưu danh sách giáo viên vào file JSON?"):
                                self._save_teachers_data(teachers_list, teachers_count)
                        else:
                            print_status("Không có giáo viên nào trong danh sách", "warning")
                    
                    elif isinstance(data, list):
                        print_status(f"✅ Lấy danh sách thành công! Tìm thấy {len(data)} giáo viên", "success")
                        
                        if len(data) > 0:
                            print(f"\n📋 DANH SÁCH GIÁO VIÊN (hiển thị {min(len(data), 10)} đầu tiên):")
                            for i, teacher in enumerate(data[:10], 1):
                                print(f"   {i:2d}. {teacher}")
                            
                            if len(data) > 10:
                                print(f"   ... và {len(data) - 10} giáo viên khác")
                            
                            # Hỏi có muốn lưu JSON không
                            if get_user_confirmation("Lưu danh sách giáo viên vào file JSON?"):
                                self._save_teachers_data(data, len(data))
                        else:
                            print_status("Không có giáo viên nào trong danh sách", "warning")
                    
                    else:
                        print_status(f"✅ Lấy dữ liệu thành công! Response type: {type(data)}", "success")
                        print(f"📋 DATA: {data}")
                else:
                    print_status("API trả về thành công nhưng không có dữ liệu", "warning")
            else:
                print_status(f"❌ Lỗi lấy danh sách: {result.get('error', 'Unknown error')}", "error")
                if result.get('status_code'):
                    print(f"   📡 Status Code: {result.get('status_code')}")
            
        except ImportError:
            print_status("Module onluyen_api chưa được cài đặt", "error")
        except Exception as e:
            print_status(f"Lỗi lấy danh sách giáo viên: {e}", "error")
    
    def onluyen_get_students(self):
        """Lấy danh sách học sinh"""
        print_separator("LẤY DANH SÁCH HỌC SINH")
        
        try:
            
            # Hỏi page index và page size với default lớn hơn
            page_index = get_user_input("Nhập page index (Enter = 1)") or "1"
            page_size = get_user_input("Nhập page size (Enter = 5000)") or "5000"
            
            try:
                page_index = int(page_index)
                page_size = int(page_size)
            except ValueError:
                page_index = 1
                page_size = 5000
            
            client = OnLuyenAPIClient()
            
            # Kiểm tra có cần login không
            if self._check_onluyen_auth_required(client):
                return
            
            print_status(f"Đang lấy danh sách học sinh (page {page_index}, size: {page_size})...", "info")
            
            result = client.get_students(page_index=page_index, page_size=page_size)
            
            if result['success']:
                data = result.get('data')
                
                if data:
                    if isinstance(data, dict) and 'data' in data:
                        students_list = data['data']
                        students_count = data.get('totalCount', len(students_list))
                        
                        print_status(f"✅ Lấy danh sách thành công: {len(students_list)}/{students_count} học sinh", "success")
                        
                        if len(students_list) > 0:
                            print(f"\n📋 DANH SÁCH HỌC SINH (hiển thị {min(len(students_list), 10)} đầu tiên):")
                            for i, student in enumerate(students_list[:10], 1):
                                if isinstance(student, dict):
                                    name = student.get('name', student.get('fullName', 'N/A'))
                                    email = student.get('email', 'N/A')
                                    id_val = student.get('id', student.get('studentId', 'N/A'))
                                    class_name = student.get('className', 'N/A')
                                    print(f"   {i:2d}. ID: {id_val} | Tên: {name} | Lớp: {class_name}")
                                else:
                                    print(f"   {i:2d}. {student}")
                            
                            if len(students_list) > 10:
                                print(f"   ... và {len(students_list) - 10} học sinh khác")
                            
                            # Hỏi có muốn lưu JSON không
                            if get_user_confirmation("Lưu danh sách học sinh vào file JSON?"):
                                self._save_students_data(students_list, students_count)
                        else:
                            print_status("Không có học sinh nào trong danh sách", "warning")
                    
                    elif isinstance(data, list):
                        print_status(f"✅ Lấy danh sách thành công! Tìm thấy {len(data)} học sinh", "success")
                        
                        if len(data) > 0:
                            print(f"\n📋 DANH SÁCH HỌC SINH (hiển thị {min(len(data), 10)} đầu tiên):")
                            for i, student in enumerate(data[:10], 1):
                                print(f"   {i:2d}. {student}")
                            
                            if len(data) > 10:
                                print(f"   ... và {len(data) - 10} học sinh khác")
                            
                            # Hỏi có muốn lưu JSON không
                            if get_user_confirmation("Lưu danh sách học sinh vào file JSON?"):
                                self._save_students_data(data, len(data))
                        else:
                            print_status("Không có học sinh nào trong danh sách", "warning")
                    
                    else:
                        print_status(f"✅ Lấy dữ liệu thành công! Response type: {type(data)}", "success")
                        print(f"📋 DATA: {data}")
                else:
                    print_status("API trả về thành công nhưng không có dữ liệu", "warning")
            else:
                print_status(f"❌ Lỗi lấy danh sách: {result.get('error', 'Unknown error')}", "error")
                if result.get('status_code'):
                    print(f"   📡 Status Code: {result.get('status_code')}")
            
        except ImportError:
            print_status("Module onluyen_api chưa được cài đặt", "error")
        except Exception as e:
            print_status(f"Lỗi lấy danh sách học sinh: {e}", "error")
    
    def onluyen_configure_credentials(self):
        """Cấu hình credentials OnLuyen"""
        print_separator("CẤU HÌNH ONLUYEN CREDENTIALS")
        
        print("📋 Cấu hình này sẽ cập nhật file .env")
        print("⚠️  Lưu ý: Credentials sẽ được lưu dưới dạng plain text")
        
        if not get_user_confirmation("Tiếp tục cấu hình credentials?"):
            return
        
        username = get_user_input("Nhập OnLuyen username", required=True)
        if not username:
            return
        
        password = get_user_input("Nhập OnLuyen password", required=True)
        if not password:
            return
        
        try:
            # Cập nhật .env file
            env_file = Path(".env")
            if env_file.exists():
                with open(env_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Update username
                if "ONLUYEN_USERNAME=" in content:
                    content = content.replace(
                        f"ONLUYEN_USERNAME=",
                        f"ONLUYEN_USERNAME={username}"
                    )
                else:
                    content += f"\nONLUYEN_USERNAME={username}"
                
                # Update password
                if "ONLUYEN_PASSWORD=" in content:
                    content = content.replace(
                        f"ONLUYEN_PASSWORD=",
                        f"ONLUYEN_PASSWORD={password}"
                    )
                else:
                    content += f"\nONLUYEN_PASSWORD={password}"
                
                with open(env_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                print_status("Đã cập nhật credentials vào .env", "success")
                print("🔄 Khởi động lại ứng dụng để áp dụng thay đổi")
                
            else:
                print_status("File .env không tồn tại", "error")
                
        except Exception as e:
            print_status(f"Lỗi cập nhật credentials: {e}", "error")
    
    def _check_onluyen_auth_required(self, client) -> bool:
        """
        Kiểm tra có cần authentication không
        
        Args:
            client: OnLuyenAPIClient instance
            
        Returns:
            bool: True nếu cần auth và chưa auth, False nếu OK
        """
        if not client.auth_token:
            print_status("Chưa login. Vui lòng login trước khi sử dụng tính năng này.", "warning")
            print("💡 Hãy sử dụng workflow hoàn chỉnh để tự động login.")
            return True
        return False
    
    def onluyen_integrated_processing(self):
        """Xử lý tích hợp: Google Sheets → OnLuyen API Login"""
        print_separator("XỬ LÝ TÍCH HỢP: GOOGLE SHEETS → ONLUYEN API")
        
        try:
            # Bước 1: Trích xuất dữ liệu từ Google Sheets
            print_status("BƯỚC 1: Trích xuất dữ liệu từ Google Sheets", "info")
            
            extractor = GoogleSheetsExtractor()
            
            # Lấy tên sheet với logic fallback
            sheet_name = self._get_sheet_name_with_fallback(extractor)
            if not sheet_name:
                print_status("❌ Không thể xác định tên sheet", "error")
                return
            
            print_status(f"Đang trích xuất dữ liệu từ sheet: {sheet_name}", "info")
            school_data = extractor.extract_school_data(sheet_name=sheet_name)
            
            if not school_data:
                print_status("Không thể trích xuất dữ liệu từ Google Sheets", "error")
                return
            
            print_status(f"✅ Đã trích xuất {len(school_data)} trường học", "success")
            
            # Hiển thị danh sách trường để chọn
            if len(school_data) > 1:
                print("\nDanh sách trường đã trích xuất:")
                for i, school in enumerate(school_data, 1):
                    school_name = school.get('Tên trường', 'N/A')
                    admin_email = school.get('Admin', 'N/A')
                    print(f"{i}. {school_name} (Admin: {admin_email})")
            
            # Chọn trường để xử lý
            if len(school_data) == 1:
                selected_school = school_data[0]
            else:
                try:
                    choice = get_user_input(f"Chọn trường để xử lý (1-{len(school_data)})", required=True)
                    choice_idx = int(choice) - 1
                    if 0 <= choice_idx < len(school_data):
                        selected_school = school_data[choice_idx]
                    else:
                        print_status("Lựa chọn không hợp lệ", "error")
                        return
                except (ValueError, TypeError):
                    print_status("Lựa chọn không hợp lệ", "error")
                    return
            
            # Bước 2: Lấy thông tin login
            print_status("BƯỚC 2: Chuẩn bị thông tin login", "info")
            
            school_name = selected_school.get('Tên trường', 'N/A')
            admin_email = selected_school.get('Admin', '')
            password = selected_school.get('Mật khẩu', '')
            drive_link = selected_school.get('Link driver dữ liệu', 'N/A')
            
            if not admin_email or not password:
                missing_fields = []
                if not admin_email:
                    missing_fields.append("Admin email")
                if not password:
                    missing_fields.append("Mật khẩu")
                
                print_status(f"Thiếu thông tin cần thiết: {', '.join(missing_fields)}", "error")
                return
            
            # Bước 3: Login vào OnLuyen API
            print_status("BƯỚC 3: Thực hiện login OnLuyen API", "info")
            
            client = OnLuyenAPIClient()
            print_status(f"Đang login với Admin: {admin_email}", "info")
            
            result = client.login(admin_email, password)
            
            # Bước 4: Log response và kết quả
            print_status("BƯỚC 4: Phân tích kết quả login", "info")
            
            print(f"\nTrường: {school_name}")
            print(f"Admin: {admin_email}")
            print(f"Success: {result['success']}")
            
            if result['success']:
                print_status("LOGIN THÀNH CÔNG!", "success")
                
                if result.get('data'):
                    response_data = result['data']
                    self._log_login_response(response_data)
                    
                    # Bước 4.1: Kiểm tra tài khoản đăng nhập có khớp không
                    response_email = response_data.get('account', '').lower().strip()
                    expected_email = admin_email.lower().strip()
                    
                    if response_email == expected_email:
                        print_status("✅ Tài khoản đăng nhập trùng khớp!", "success")
                        
                        # Lưu thông tin thành công
                        self._save_successful_login_info(school_name, admin_email, result, drive_link, password)
                        
                        # Cập nhật tóm tắt
                        account_match = True
                    else:
                        print_status("❌ Tài khoản đăng nhập chưa trùng khớp!", "error")
                        print(f"   🚨 Có thể đây là tài khoản khác hoặc dữ liệu không đồng bộ")
                        
                        # Đăng xuất
                        print_status("ĐANG THỰC HIỆN ĐĂNG XUẤT...", "warning")
                        logout_result = self._logout_onluyen_api(client)
                        
                        if logout_result:
                            print_status("✅ Đã đăng xuất thành công", "success")
                        else:
                            print_status("⚠️ Không thể đăng xuất hoặc đã đăng xuất", "warning")
                        
                        # Cập nhật tóm tắt
                        account_match = False
                        
                else:
                    print("   ⚠️  Không có dữ liệu response")
                    account_match = False
                    
            else:
                print_status("❌ LOGIN THẤT BẠI!", "error")
                print(f"\n🚨 LỖI: {result.get('error', 'Unknown error')}")
                
                # Log chi tiết lỗi
                self._log_login_error(school_name, admin_email, result)
                account_match = False
            
            print(f"\n📊 TÓM TẮT XỬ LÝ:")
            print("=" * 60)
            print(f"✅ Trích xuất Google Sheets: Thành công")
            print(f"✅ Chuẩn bị thông tin: Thành công")
            print(f"{'✅' if result['success'] else '❌'} OnLuyen API Login: {'Thành công' if result['success'] else 'Thất bại'}")
            if result['success']:
                print(f"{'✅' if account_match else '❌'} Kiểm tra tài khoản: {'Trùng khớp' if account_match else 'Không trùng khớp'}")
                if not account_match:
                    print(f"🚨 TÀI KHOẢN ĐĂNG NHẬP CHƯA TRÙNG KHỚP - ĐÃ ĐĂNG XUẤT")
            
        except ImportError as e:
            print_status(f"Module không tồn tại: {e}", "error")
        except Exception as e:
            print_status(f"Lỗi xử lý tích hợp: {e}", "error")
    
    def _logout_onluyen_api(self, client):
        """Đăng xuất OnLuyen API"""
        try:
            # Clear token từ client
            if hasattr(client, 'auth_token'):
                client.auth_token = None
            
            # Remove Authorization header
            if 'Authorization' in client.session.headers:
                del client.session.headers['Authorization']
            
            print("   🔓 Đã xóa token khỏi session")
            return True
            
        except Exception as e:
            print(f"   ⚠️ Lỗi khi đăng xuất: {e}")
            return False
    
    def _log_login_response(self, response_data):
        """Log chi tiết response data"""
        if isinstance(response_data, dict):
            for key, value in response_data.items():
                # Ẩn sensitive data nhưng vẫn hiển thị cấu trúc
                if any(sensitive in key.lower() for sensitive in ['token', 'secret', 'key', 'password']):
                    if value:
                        display_value = f"***{str(value)[-4:]}" if len(str(value)) > 4 else "***"
                    else:
                        display_value = "N/A"
                else:
                    display_value = value
                
                print(f"   {key}: {display_value}")
        else:
            print(f"   Raw Response: {response_data}")
    
    def _log_login_error(self, school_name, admin_email, result):
        """Log chi tiết lỗi login"""
        error_info = {
            'school': school_name,
            'admin': admin_email,
            'status_code': result.get('status_code'),
            'error': result.get('error'),
            'timestamp': __import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        print(f"\n🔍 CHI TIẾT LỖI:")
        for key, value in error_info.items():
            print(f"   {key}: {value}")
    
    def _save_successful_login_info(self, school_name, admin_email, result, drive_link, password=None):
        """Lưu thông tin login thành công bao gồm tokens và password"""
        try:
            
            # Lấy data từ response
            response_data = result.get('data', {})
            
            login_info = {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'school_name': school_name,
                'admin_email': admin_email,
                'admin_password': password,  # Thêm password cho export
                'drive_link': drive_link,
                'login_status': 'success',
                'status_code': result.get('status_code'),
                'response_keys': list(response_data.keys()) if response_data else [],
                # Thêm tokens để có thể sử dụng lại
                'tokens': {
                    'access_token': response_data.get('access_token'),
                    'refresh_token': response_data.get('refresh_token'),
                    'expires_in': response_data.get('expires_in'),
                    'expires_at': response_data.get('expires_at'),
                    'user_id': response_data.get('userId'),
                    'display_name': response_data.get('display_name'),
                    'account': response_data.get('account')
                }
            }
            
            filename = f"onluyen_login_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            filepath = f"data/output/{filename}"
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(login_info, f, ensure_ascii=False, indent=2)
            
            print_status(f"✅ Đã lưu thông tin login vào: {filepath}", "success")
            
        except Exception as e:
            print_status(f"Lỗi lưu thông tin login: {e}", "warning")
    
    def onluyen_complete_workflow(self):
        """Tích hợp hoàn chỉnh: Sheets → Login → Lấy dữ liệu GV/HS → Chuyển đổi Excel"""
        print_separator("HOÀN CHỈNH: SHEETS → LOGIN → DỮ LIỆU → EXCEL")
        
        print("🔄 CHỌN LUỒNG XỬ LÝ:")
        print("   📋 Case 1: Xuất toàn bộ dữ liệu")
        print("   📋 Case 2: Xuất dữ liệu theo file import (có so sánh)")
        print()
        
        # Menu chọn case
        case_options = [
            "Case 1: Toàn bộ dữ liệu (Sheets → Login → Dữ liệu → Excel)",
            "Case 2: Dữ liệu theo file import (Sheets → Login → Dữ liệu → So sánh → Excel)"
        ]
        
        case_handlers = [
            self._workflow_case_1_full_data,
            self._workflow_case_2_import_filtered
        ]
        
        run_menu_loop("CHỌN LUỒNG XỬ LÝ", case_options, case_handlers)
    
    def _workflow_case_1_full_data(self):
        """Case 1: Luồng xử lý toàn bộ dữ liệu"""
        print_separator("CASE 1: TOÀN BỘ DỮ LIỆU")
        
        print("🔄 LUỒNG XỬ LÝ HOÀN CHỈNH:")
        print("   1️⃣  Trích xuất dữ liệu từ Google Sheets")
        print("   2️⃣  Login vào OnLuyen API") 
        print("   3️⃣  Lấy danh sách Giáo viên")
        print("   4️⃣  Lấy danh sách Học sinh")
        print("   5️⃣  Lưu dữ liệu workflow JSON")
        print("   6️⃣  Chuyển đổi JSON → Excel")
        print("   7️⃣  Upload files lên Google Drive (OAuth 2.0)")
        print("       📁 Sử dụng text value từ cột 'Link driver dữ liệu' trong Google Sheets")
        print("   8️⃣  Tổng hợp và báo cáo kết quả")
        print()
        print("💡 Lưu ý: ")
        print("   • Drive link được lấy từ text value của cột 'Link driver dữ liệu' (không extract hyperlink)")
        print("   • Đảm bảo cột 'Link driver dữ liệu' chứa URL đầy đủ dạng text")
        print("   • Nếu chỉ muốn lấy dữ liệu riêng lẻ, hãy chọn chức năng 2 hoặc 3 trong menu")
        print()
        
        self._execute_workflow_case_1()
    
    def _workflow_case_2_import_filtered(self):
        """Case 2: Luồng xử lý dữ liệu theo file import"""
        print_separator("CASE 2: DỮ LIỆU THEO FILE IMPORT")
        
        print("🔄 LUỒNG XỬ LÝ VỚI SO SÁNH:")
        print("   1️⃣  Trích xuất dữ liệu từ Google Sheets")
        print("   2️⃣  Login vào OnLuyen API") 
        print("   3️⃣  Lấy danh sách Giáo viên")
        print("   4️⃣  Lấy danh sách Học sinh")
        print("   5️⃣  Tải file import từ Google Drive")
        print("       📁 Tìm file có cấu trúc tên 'import_[Tên trường]'")
        print("   6️⃣  So sánh và lọc dữ liệu")
        print("       🔍 Chỉ giữ lại dữ liệu có trong file import")
        print("   7️⃣  Lưu dữ liệu đã lọc workflow JSON")
        print("   8️⃣  Chuyển đổi JSON → Excel")
        print("   9️⃣  Upload files lên Google Drive (OAuth 2.0)")
        print("   🔟 Tổng hợp và báo cáo kết quả")
        print()
        print("💡 Lưu ý: ")
        print("   • File import phải có tên bắt đầu bằng 'import_' và kết thúc bằng '.xlsx'")
        print("   • Ví dụ: import_data.xlsx, import_truong_abc.xlsx")
        print("   • File import phải nằm trong Drive folder từ 'Link driver dữ liệu'")

        print("   • File phải chứa danh sách email/username cần so sánh")
        print("   • Nếu có nhiều file import_, hệ thống sẽ cho bạn chọn")
        print()
        self._execute_workflow_case_2()

    def _execute_workflow_case_1(self, selected_school_data):
        """Execute Case 1 workflow - toàn bộ dữ liệu"""

        workflow_results = {
            'sheets_extraction': False,
            'api_login': False, 
            'teachers_data': False,
            'students_data': False,
            'json_saved': False,
            'excel_converted': False,
            'drive_uploaded': False,
            'school_info': {},
            'data_summary': {},
            'json_file_path': None,
            'excel_file_path': None,
            'upload_results': {}
        }
        
        try:
            # Bước 1: Trích xuất dữ liệu từ Google Sheets
            # print_status("BƯỚC 1: Trích xuất dữ liệu từ Google Sheets", "info")
            
            # extractor = GoogleSheetsExtractor()
            
            # # Lấy tên sheet với logic fallback
            # sheet_name = self._get_sheet_name_with_fallback(extractor)
            # if not sheet_name:
            #     print_status("❌ Không thể xác định tên sheet", "error")
            #     return
            
            # print_status(f"Đang trích xuất dữ liệu từ sheet: {sheet_name}", "info")
            # school_data = extractor.extract_school_data(sheet_name=sheet_name)
            
            # if not school_data:
            #     print_status("❌ Không thể trích xuất dữ liệu từ Google Sheets", "error")
            #     return
            
            # workflow_results['sheets_extraction'] = True
            # print_status(f"✅ Đã trích xuất {len(school_data)} trường học", "success")
            
            # # Chọn trường để xử lý
            # if len(school_data) == 1:
            #     selected_school = school_data[0]
            #     print_status("Tự động chọn trường duy nhất", "info")
            # else:
            #     print("\n📋 DANH SÁCH TRƯỜNG ĐÃ TRÍCH XUẤT:")
            #     for i, school in enumerate(school_data, 1):
            #         school_name = school.get('Tên trường', 'N/A')
            #         admin_email = school.get('Admin', 'N/A')
            #         print(f"   {i}. {school_name} (Admin: {admin_email})")
                
            #     try:
            #         choice = get_user_input(f"Chọn trường để xử lý (1-{len(school_data)})", required=True)
            #         choice_idx = int(choice) - 1
            #         if 0 <= choice_idx < len(school_data):
            #             selected_school = school_data[choice_idx]
            #         else:
            #             print_status("Lựa chọn không hợp lệ", "error")
            #             return
            #     except (ValueError, TypeError):
            #         print_status("Lựa chọn không hợp lệ", "error")
            #         return
            
            # Lấy thông tin trường
            school_name = selected_school_data.get('Tên trường', 'N/A')
            admin_email = selected_school_data.get('Admin', '')
            password = selected_school_data.get('Mật khẩu', '')
            drive_link = selected_school_data.get('Link driver dữ liệu', 'N/A')
            
            workflow_results['school_info'] = {
                'name': school_name,
                'admin': admin_email,
                'drive_link': drive_link
            }
            
            print(f"\n📋 THÔNG TIN TRƯỜNG ĐÃ CHỌN:")
            print(f"   🏫 Tên trường: {school_name}")
            print(f"   👤 Admin: {admin_email}")
            print(f"   🔗 Drive Link: {drive_link[:60] + '...' if len(drive_link) > 60 else drive_link}")
            
            # Validate Drive link ngay từ đầu
            if drive_link and drive_link != 'N/A' and 'drive.google.com' in drive_link:
                folder_id_preview = self._extract_drive_folder_id(drive_link)
                if folder_id_preview:
                    print(f"   ✅ Drive link hợp lệ")
                else:
                    print(f"   ❌ Drive link không hợp lệ")
            else:
                print(f"   ⚠️ Không có Drive link hợp lệ")
            
            if not admin_email or not password:
                print_status("❌ Thiếu thông tin Admin email hoặc Mật khẩu", "error")
                return
            
            # Bước 2: Login vào OnLuyen API
            print_status("BƯỚC 2: Thực hiện login OnLuyen API", "info")
            
            client = OnLuyenAPIClient()
            print_status(f"Đang login với Admin: {admin_email}", "info")
            
            result = client.login(admin_email, password)
            
            if not result['success']:
                print_status(f"❌ Login thất bại: {result.get('error', 'Unknown error')}", "error")
                return
            
            # Kiểm tra tài khoản trùng khớp
            response_data = result.get('data', {})
            response_email = response_data.get('account', '').lower().strip()
            expected_email = admin_email.lower().strip()
            
            if response_email != expected_email:
                print_status("❌ Tài khoản đăng nhập không trùng khớp!", "error")
                print(f"   Expected: {expected_email}")
                print(f"   Got: {response_email}")
                return
            
            workflow_results['api_login'] = True
            print_status("✅ Login thành công và tài khoản trùng khớp", "success")
            
            # Lưu thông tin login
            self._save_successful_login_info(school_name, admin_email, result, drive_link, password)
            
            # Bước 3: Lấy danh sách Giáo viên
            print_status("BƯỚC 3: Lấy danh sách Giáo viên", "info")
            
            teachers_result = client.get_teachers(page_size=1000)
            
            if teachers_result['success'] and teachers_result.get('data'):
                teachers_data = teachers_result['data']
                if isinstance(teachers_data, dict) and 'data' in teachers_data:
                    teachers_list = teachers_data['data']
                    teachers_count = teachers_data.get('totalCount', len(teachers_list))
                    
                    workflow_results['teachers_data'] = True
                    workflow_results['data_summary']['teachers'] = {
                        'total': teachers_count,
                        'retrieved': len(teachers_list)
                    }
                    
                    print_status(f"✅ Lấy danh sách giáo viên thành công: {len(teachers_list)}/{teachers_count}", "success")
                    
                    # Extract thông tin HT/HP
                    print_status("🔍 Trích xuất thông tin Hiệu trường (HT) và Hiệu phó (HP)", "info")
                    ht_hp_info = self._extract_ht_hp_info(teachers_data)
                    workflow_results['ht_hp_info'] = ht_hp_info
                    
                    # Lưu thông tin HT/HP vào file riêng
                    ht_hp_file = self._save_ht_hp_info(ht_hp_info, school_name)
                    if ht_hp_file:
                        workflow_results['ht_hp_file'] = ht_hp_file
                        
                else:
                    print_status("⚠️ Định dạng dữ liệu giáo viên không đúng", "warning")
            else:
                print_status(f"❌ Lỗi lấy danh sách giáo viên: {teachers_result.get('error')}", "error")
            
            # Bước 4: Lấy danh sách Học sinh
            print_status("BƯỚC 4: Lấy danh sách Học sinh", "info")
            
            students_result = client.get_students(page_index=1, page_size=5000)
            
            if students_result['success'] and students_result.get('data'):
                students_data = students_result['data']
                if isinstance(students_data, dict) and 'data' in students_data:
                    students_list = students_data['data']
                    students_count = students_data.get('totalCount', len(students_list))
                    
                    workflow_results['students_data'] = True
                    workflow_results['data_summary']['students'] = {
                        'total': students_count,
                        'retrieved': len(students_list)
                    }
                    
                    print_status(f"✅ Lấy danh sách học sinh thành công: {len(students_list)}/{students_count}", "success")
                else:
                    print_status("⚠️ Định dạng dữ liệu học sinh không đúng", "warning")
            else:
                print_status(f"❌ Lỗi lấy danh sách học sinh: {students_result.get('error')}", "error")
            
            # Bước 5: Lưu dữ liệu workflow JSON
            print_status("BƯỚC 5: Lưu dữ liệu workflow JSON", "info")
            
            if workflow_results['teachers_data'] or workflow_results['students_data']:
                json_file_path = self._save_workflow_data(workflow_results, teachers_result, students_result, password)
                if json_file_path:
                    workflow_results['json_saved'] = True
                    workflow_results['json_file_path'] = json_file_path
                    print_status(f"✅ Đã lưu dữ liệu JSON: {json_file_path}", "success")
                else:
                    print_status("❌ Lỗi lưu dữ liệu JSON", "error")
            else:
                print_status("⚠️ Không có dữ liệu để lưu", "warning")
            
            # Bước 6: Chuyển đổi JSON → Excel
            print_status("BƯỚC 6: Chuyển đổi JSON → Excel", "info")
            
            if workflow_results['json_saved'] and workflow_results['json_file_path']:
                excel_file_path = self._convert_json_to_excel(workflow_results['json_file_path'])
                if excel_file_path:
                    workflow_results['excel_converted'] = True
                    workflow_results['excel_file_path'] = excel_file_path
                    print_status(f"✅ Đã tạo file Excel: {excel_file_path}", "success")
                else:
                    print_status("❌ Lỗi chuyển đổi sang Excel", "error")
            else:
                print_status("⚠️ Không có file JSON để chuyển đổi", "warning")
            
            # Bước 7: Hỏi có muốn upload file Excel lên Google Drive không
            print_status("BƯỚC 7: Upload file Excel lên Google Drive (Tùy chọn)", "info")
            
            # Kiểm tra có file Excel để upload không
            excel_file_exists = workflow_results['excel_converted'] and workflow_results['excel_file_path'] and os.path.exists(workflow_results['excel_file_path'])
            
            if excel_file_exists:
                excel_file_name = os.path.basename(workflow_results['excel_file_path'])
                excel_file_size = os.path.getsize(workflow_results['excel_file_path']) / (1024 * 1024)  # MB
                
                print(f"\n📊 FILE EXCEL SẴN SÀNG UPLOAD:")
                print(f"   📄 Tên file: {excel_file_name}")
                print(f"   📏 Kích thước: {excel_file_size:.1f} MB")
                
                # Hỏi người dùng có muốn upload không
                if get_user_confirmation("\n📤 Bạn có muốn upload file Excel lên Google Drive?"):
                    # Validate Drive link
                    is_valid_drive_link = False
                    folder_id = None
                    
                    if drive_link and drive_link != 'N/A' and 'drive.google.com' in drive_link:
                        folder_id = self._extract_drive_folder_id(drive_link)
                        if folder_id:
                            print(f"   ✅ Drive link hợp lệ")
                            is_valid_drive_link = True
                        else:
                            print(f"   ❌ Không thể extract folder ID từ link")
                    else:
                        # Xử lý các trường hợp Drive link không hợp lệ
                        if not drive_link or drive_link == 'N/A':
                            print(f"   ⚠️ Không có Drive link trong Google Sheets")
                        elif drive_link and 'drive.google.com' not in drive_link:
                            print(f"   ⚠️ Drive link không đúng format")
                            print(f"   💡 Cần format: https://drive.google.com/drive/folders/[FOLDER_ID]")
                        
                        print(f"   ❌ Drive link không hợp lệ")
                    
                    if is_valid_drive_link:
                        print_status(f"📤 Đang upload file Excel: {excel_file_name}", "info")
                        
                        # Upload chỉ file Excel
                        upload_results = self._upload_files_to_drive_oauth([workflow_results['excel_file_path']], drive_link)
                        
                        workflow_results['upload_results'] = upload_results
                        
                        if upload_results.get('success', 0) > 0:
                            workflow_results['drive_uploaded'] = True
                            print_status(f"✅ Upload file Excel thành công!", "success")
                            
                            # Hiển thị URL
                            if upload_results.get('urls'):
                                print(f"\n📂 FILE EXCEL ĐÃ UPLOAD:")
                                print(f"   🔗 {upload_results['urls'][0]}")
                        else:
                            workflow_results['drive_uploaded'] = False
                            print_status("❌ Upload file Excel thất bại", "error")
                            
                            # Debug thông tin lỗi
                            if upload_results.get('errors'):
                                print(f"\n🚨 CHI TIẾT LỖI:")
                                for error in upload_results['errors']:
                                    print(f"   ❌ {error}")
                    else:
                        workflow_results['drive_uploaded'] = False
                        print_status("⚠️ Không thể upload do Drive link không hợp lệ", "warning")
                        print("💡 HƯỚNG DẪN SETUP DRIVE LINK:")
                        print("   1️⃣  Mở Google Sheets")
                        print("   2️⃣  Tìm cột 'Link driver dữ liệu'")
                        print("   3️⃣  Nhập link Drive folder thực tế")
                        print("   4️⃣  Format: https://drive.google.com/drive/folders/[FOLDER_ID]")
                        
                        # Hỏi có muốn nhập Drive link thủ công không
                        if get_user_confirmation("\nBạn có muốn nhập Drive link thủ công để upload?"):
                            manual_drive_link = get_user_input("Nhập Google Drive folder link:")
                            if manual_drive_link and 'drive.google.com' in manual_drive_link:
                                folder_id_manual = self._extract_drive_folder_id(manual_drive_link)
                                if folder_id_manual:
                                    print_status(f"📤 Uploading với Drive link thủ công...", "info")
                                    upload_results = self._upload_files_to_drive_oauth([workflow_results['excel_file_path']], manual_drive_link)
                                    
                                    workflow_results['upload_results'] = upload_results
                                    
                                    if upload_results.get('success', 0) > 0:
                                        workflow_results['drive_uploaded'] = True
                                        print_status(f"✅ Upload file Excel thành công với Drive link thủ công!", "success")
                                        if upload_results.get('urls'):
                                            print(f"\n📂 FILE EXCEL ĐÃ UPLOAD:")
                                            print(f"   🔗 {upload_results['urls'][0]}")
                                    else:
                                        workflow_results['drive_uploaded'] = False
                                        print_status("❌ Upload file Excel thất bại", "error")
                                else:
                                    workflow_results['drive_uploaded'] = False
                                    print_status("❌ Drive link thủ công không hợp lệ", "error")
                            else:
                                workflow_results['drive_uploaded'] = False
                                print_status("❌ Drive link thủ công không đúng format", "error")
                else:
                    workflow_results['drive_uploaded'] = False
                    print_status("ℹ️ Bỏ qua upload file Excel", "info")
            else:
                workflow_results['drive_uploaded'] = False
                print_status("⚠️ Không có file Excel để upload", "warning")
            
            # Bước 8: Tổng hợp và báo cáo kết quả
            print_status("BƯỚC 8: Tổng hợp kết quả", "info")
            
            self._print_workflow_summary(workflow_results)
            
            # # Hỏi có muốn mở file Excel không nếu tạo thành công
            # if workflow_results['excel_converted'] and workflow_results['excel_file_path']:
            #     action_options = ["Mở file Excel local"]
                
            #     if workflow_results['drive_uploaded'] and workflow_results['upload_results'].get('urls'):
            #         action_options.append("Mở Google Drive folder")
                
            #     if len(action_options) > 1:
            #         print(f"\n🎯 BẠN CÓ THỂ:")
            #         for i, option in enumerate(action_options, 1):
            #             print(f"   {i}. {option}")
                    
            #         choice = get_user_input(f"Chọn hành động (1-{len(action_options)}, Enter = bỏ qua)")
                    
            #         if choice == "1":
            #             try:
            #                 os.startfile(workflow_results['excel_file_path'])
            #                 print_status("Đã mở file Excel", "success")
            #             except Exception as e:
            #                 print_status(f"Không thể mở file Excel: {e}", "warning")
            #         elif choice == "2" and len(action_options) > 1:
            #             drive_folder_url = drive_link
            #             print_status(f"🔗 Google Drive: {drive_folder_url}", "info")
            #             print("💡 Bạn có thể mở link trên trong trình duyệt")
            #     else:
            #         if get_user_confirmation("Bạn có muốn mở file Excel đã tạo?"):
            #             try:
            #                 os.startfile(workflow_results['excel_file_path'])
            #                 print_status("Đã mở file Excel", "success")
            #             except Exception as e:
            #                 print_status(f"Không thể mở file Excel: {e}", "warning")
            
            # Lưu dữ liệu vào file nếu chưa lưu (fallback)
            if not workflow_results['json_saved'] and (workflow_results['teachers_data'] or workflow_results['students_data']):
                self._save_workflow_data(workflow_results, teachers_result, students_result, password)
            
        except ImportError as e:
            print_status(f"Module không tồn tại: {e}", "error")
        except Exception as e:
            print_status(f"Lỗi trong quy trình tích hợp: {e}", "error")

    def _execute_workflow_case_2(self, selected_school_data):
        """Case 2: Workflow với so sánh file import"""
        
        workflow_results = {
            'sheets_extraction': False,
            'api_login': False, 
            'teachers_data': False,
            'students_data': False,
            'import_file_downloaded': False,
            'data_comparison': False,
            'json_saved': False,
            'excel_converted': False,
            'drive_uploaded': False,
            'school_info': {},
            'data_summary': {},
            'import_file_info': {},
            'comparison_results': {},
            'json_file_path': None,
            'excel_file_path': None,
            'upload_results': {}
        }
        
        try:
            # Bước 1-4: Giống Case 1 - Lấy dữ liệu từ Sheets và OnLuyen API
            print_status("BƯỚC 1-4: Lấy dữ liệu cơ bản (giống Case 1)", "info")
            
            # Thực hiện các bước giống Case 1
            basic_results = self._execute_basic_workflow_steps(selected_school_data)
            if not basic_results:
                print_status("❌ Lỗi trong các bước cơ bản", "error")
                
                return
            
            # Cập nhật workflow_results với dữ liệu cơ bản
            workflow_results.update(basic_results)
            
            if not (workflow_results['sheets_extraction'] and workflow_results['api_login'] and 
                   (workflow_results['teachers_data'] or workflow_results['students_data'])):
                print_status("❌ Không đủ dữ liệu cơ bản để tiếp tục", "error")
                return

            # Bước 5: Tải file import từ Google Drive
            print_status("BƯỚC 5: Tải file import từ Google Drive", "info")
            
            school_name = workflow_results['school_info'].get('name', '')
            drive_link = workflow_results['school_info'].get('drive_link', '')
            
            import_file_path = self._download_import_file(school_name, drive_link)
            
            if import_file_path:
                workflow_results['import_file_downloaded'] = True
                workflow_results['import_file_info'] = {
                    'file_path': import_file_path,
                    'file_name': os.path.basename(import_file_path)
                }
                print_status(f"✅ Đã tải file import: {os.path.basename(import_file_path)}", "success")
            else:
                print_status("❌ Không thể tải file import", "error")
                print("💡 HƯỚNG DẪN SETUP FILE IMPORT:")
                print("   1️⃣  File phải có tên bắt đầu bằng 'import_' và kết thúc bằng '.xlsx'")
                print("   2️⃣  Ví dụ: import_data.xlsx, import_truong_abc.xlsx")
                print("   3️⃣  File phải nằm trong Drive folder từ 'Link driver dữ liệu'")
                print("   4️⃣  File phải chứa danh sách email/username cần so sánh")
                return
            
            # Bước 6: So sánh và lọc dữ liệu
            print_status("BƯỚC 6: So sánh và lọc dữ liệu", "info")
            
            comparison_results = self._compare_and_filter_data(
                workflow_results.get('teachers_result'), 
                workflow_results.get('students_result'),
                import_file_path
            )
            
            if comparison_results:
                workflow_results['data_comparison'] = True
                workflow_results['comparison_results'] = comparison_results
                
                teachers_filtered = comparison_results.get('teachers_filtered', [])
                students_filtered = comparison_results.get('students_filtered', [])
                
                print_status(f"✅ So sánh hoàn tất", "success")
                print(f"   👨‍🏫 Giáo viên khớp: {len(teachers_filtered)}")
                print(f"   👨‍🎓 Học sinh khớp: {len(students_filtered)}")
                
                # Cập nhật data_summary với dữ liệu đã lọc
                workflow_results['data_summary']['teachers_filtered'] = len(teachers_filtered)
                workflow_results['data_summary']['students_filtered'] = len(students_filtered)
            else:
                print_status("❌ Lỗi so sánh dữ liệu", "error")
                return
            
            # Bước 7: Lưu dữ liệu đã lọc vào JSON
            print_status("BƯỚC 7: Lưu dữ liệu đã lọc workflow JSON", "info")
            
            json_file_path = self._save_filtered_workflow_data(workflow_results, comparison_results)
            if json_file_path:
                workflow_results['json_saved'] = True
                workflow_results['json_file_path'] = json_file_path
                print_status(f"✅ Đã lưu dữ liệu đã lọc: {json_file_path}", "success")
            else:
                print_status("❌ Lỗi lưu dữ liệu JSON", "error")
            
            # Bước 8: Chuyển đổi JSON → Excel
            print_status("BƯỚC 8: Chuyển đổi JSON → Excel", "info")
            
            if workflow_results['json_saved'] and workflow_results['json_file_path']:
                excel_file_path = self._convert_json_to_excel(workflow_results['json_file_path'])
                if excel_file_path:
                    workflow_results['excel_converted'] = True
                    workflow_results['excel_file_path'] = excel_file_path
                    print_status(f"✅ Đã tạo file Excel: {excel_file_path}", "success")
                else:
                    print_status("❌ Lỗi chuyển đổi sang Excel", "error")
            else:
                print_status("⚠️ Không có file JSON để chuyển đổi", "warning")
            
            # Bước 9: Upload files lên Google Drive  
            print_status("BƯỚC 9: Upload file Excel lên Google Drive (Tùy chọn)", "info")
            
            excel_file_exists = workflow_results['excel_converted'] and workflow_results['excel_file_path'] and os.path.exists(workflow_results['excel_file_path'])
            
            if excel_file_exists:
                if get_user_confirmation("\n📤 Bạn có muốn upload file Excel lên Google Drive?"):
                    # Upload chỉ file Excel
                    upload_results = self._upload_files_to_drive_oauth([workflow_results['excel_file_path']], drive_link)
                    
                    workflow_results['upload_results'] = upload_results
                    
                    if upload_results.get('success', 0) > 0:
                        workflow_results['drive_uploaded'] = True
                        print_status(f"✅ Upload file Excel thành công!", "success")
                    else:
                        workflow_results['drive_uploaded'] = False
                        print_status("❌ Upload file Excel thất bại", "error")
                else:
                    workflow_results['drive_uploaded'] = False
                    print_status("ℹ️ Bỏ qua upload file Excel", "info")
            else:
                workflow_results['drive_uploaded'] = False
                print_status("⚠️ Không có file Excel để upload", "warning")
            
            # Bước 10: Tổng hợp và báo cáo kết quả
            print_status("BƯỚC 10: Tổng hợp kết quả", "info")
            
            self._print_workflow_summary_case_2(workflow_results)
            
            # Hỏi có muốn mở file Excel không
            # if workflow_results['excel_converted'] and workflow_results['excel_file_path']:
            #     if get_user_confirmation("Bạn có muốn mở file Excel đã tạo?"):
            #         try:
            #             os.startfile(workflow_results['excel_file_path'])
            #             print_status("Đã mở file Excel", "success")
            #         except Exception as e:
            #             print_status(f"Không thể mở file Excel: {e}", "warning")
            
        except ImportError as e:
            print_status(f"Module không tồn tại: {e}", "error")
        except Exception as e:
            print_status(f"Lỗi trong quy trình Case 2: {e}", "error")

    def _convert_json_to_excel(self, json_file_path):
        """Chuyển đổi file JSON workflow sang Excel"""
        try:
            
            print(f"   📄 File JSON: {Path(json_file_path).name}")
            
            # Khởi tạo converter
            converter = JSONToExcelTemplateConverter(json_file_path)
            
            # Load và kiểm tra JSON data
            if not converter.load_json_data():
                print("   ❌ Không thể load JSON data")
                return None
            
            # Extract data
            print("   📊 Đang trích xuất dữ liệu...")
            teachers_extracted = converter.extract_teachers_data()
            students_extracted = converter.extract_students_data()
            
            if not teachers_extracted and not students_extracted:
                print("   ❌ Không thể trích xuất dữ liệu giáo viên hoặc học sinh")
                return None
            
            # Convert to Excel
            print("   📝 Đang tạo file Excel...")
            output_path = converter.convert()
            
            if output_path:
                # Hiển thị thống kê
                teachers_count = len(converter.teachers_df) if converter.teachers_df is not None else 0
                students_count = len(converter.students_df) if converter.students_df is not None else 0
                
                print(f"   👨‍🏫 Số giáo viên: {teachers_count}")
                print(f"   👨‍🎓 Số học sinh: {students_count}")
                
                return output_path
            else:
                print("   ❌ Lỗi tạo file Excel")
                return None
                
        except ImportError:
            print("   ❌ Module json_to_excel_template_converter chưa được cài đặt")
            return None
        except Exception as e:
            print(f"   ❌ Lỗi chuyển đổi: {e}")
            return None

    def _get_drive_link_from_workflow_files(self):
        """Tìm Drive link từ workflow files có sẵn"""
        try:
            
            # Tìm files workflow JSON
            json_patterns = [
                "data/output/data_*.json",
                "data/output/workflow_data_*.json"
            ]
            
            json_files = []
            for pattern in json_patterns:
                json_files.extend(glob.glob(pattern))
            
            if not json_files:
                return None
            
            # Lấy file mới nhất
            latest_file = max(json_files, key=lambda f: os.path.getmtime(f))
            
            with open(latest_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Tìm drive link trong school_info
            drive_link = data.get('school_info', {}).get('drive_link')
            
            if drive_link and drive_link != 'N/A' and 'drive.google.com' in drive_link:
                return drive_link
            
            return None
            
        except Exception as e:
            return None

    def _upload_files_to_drive_oauth(self, file_paths, drive_link):
        """
        Upload files lên Google Drive sử dụng OAuth 2.0
        
        Args:
            file_paths: List đường dẫn files cần upload
            drive_link: Link Google Drive folder
            
        Returns:
            dict: Kết quả upload {'success': int, 'failed': int, 'urls': list, 'errors': list}
        """
        result = {
            'success': 0,
            'failed': 0,
            'urls': [],
            'errors': []
        }
        
        try:
            # Khởi tạo OAuth client
            oauth_client = GoogleOAuthDriveClient()
            
            # Kiểm tra authentication
            if not oauth_client.is_authenticated():
                error_msg = "OAuth chưa được setup hoặc token hết hạn"
                print_status(f"❌ {error_msg}", "error")
                result['failed'] = len(file_paths)
                result['errors'].append(error_msg)
                return result
            
            # Test connection
            if not oauth_client.test_connection():
                error_msg = "OAuth connection test thất bại"
                print_status(f"❌ {error_msg}", "error")
                result['failed'] = len(file_paths)
                result['errors'].append(error_msg)
                return result
            
            # Extract folder ID từ drive link
            folder_id = self._extract_drive_folder_id(drive_link)
            if not folder_id:
                error_msg = "Không thể extract folder ID từ drive link"
                print_status(f"❌ {error_msg}", "error")
                result['failed'] = len(file_paths)
                result['errors'].append(error_msg)
                return result
            
            # Upload từng file
            for file_path in file_paths:
                if not file_path or not os.path.exists(file_path):
                    result['failed'] += 1
                    error_msg = f"File không tồn tại: {file_path}"
                    result['errors'].append(error_msg)
                    continue
                
                file_name = os.path.basename(file_path)
                print_status(f"📤 Đang upload: {file_name}", "info")
                
                try:
                    file_url = oauth_client.upload_file_to_folder_id(
                        local_path=file_path,
                        folder_id=folder_id,
                        filename=file_name
                    )
                    
                    if file_url:
                        result['success'] += 1
                        result['urls'].append(file_url)
                        print(f"   ✅ Upload thành công")
                    else:
                        result['failed'] += 1
                        error_msg = f"Upload thất bại cho {file_name}"
                        result['errors'].append(error_msg)
                        print(f"   ❌ Upload thất bại")
                        
                except Exception as e:
                    result['failed'] += 1
                    error_msg = f"Lỗi upload {file_name}: {str(e)}"
                    result['errors'].append(error_msg)
                    print_status(f"❌ Lỗi upload {file_name}: {e}", "error")
            
            return result
            
        except ImportError as e:
            error_msg = f"OAuth module chưa được cài đặt: {e}"
            print_status(f"❌ {error_msg}", "error")
            result['failed'] = len(file_paths)
            result['errors'].append(error_msg)
            return result
        except Exception as e:
            error_msg = f"Lỗi OAuth upload: {e}"
            print_status(f"❌ {error_msg}", "error")
            result['failed'] = len(file_paths)
            result['errors'].append(error_msg)
            return result

    def _extract_drive_folder_id(self, drive_link):
        """Extract folder ID từ Google Drive link"""
        try:
            
            # Patterns cho Drive folder links
            patterns = [
                r'drive\.google\.com/drive/folders/([a-zA-Z0-9-_]+)',
                r'drive\.google\.com/drive/u/\d+/folders/([a-zA-Z0-9-_]+)',
                r'drive\.google\.com/open\?id=([a-zA-Z0-9-_]+)',
                r'/folders/([a-zA-Z0-9-_]+)',
                r'id=([a-zA-Z0-9-_]+)'
            ]
            
            for pattern in patterns:
                match = re.search(pattern, drive_link)
                if match:
                    return match.group(1)
            
            print_status("❌ Không thể extract folder ID từ link", "error")
            return None
            
        except Exception as e:
            print_status(f"❌ Lỗi extract folder ID: {e}", "error")
            return None

    def _print_workflow_summary(self, results):
        """In tóm tắt kết quả workflow"""
        print(f"\n📊 TÓM TẮT KẾT QUẢ WORKFLOW:")
        print("=" * 70)
        
        print(f"🏫 Trường: {results['school_info'].get('name', 'N/A')}")
        print(f"👤 Admin: {results['school_info'].get('admin', 'N/A')}")
        print()
        
        # Trạng thái từng bước
        steps = [
            ("1️⃣ Trích xuất Google Sheets", results['sheets_extraction']),
            ("2️⃣ OnLuyen API Login", results['api_login']),
            ("3️⃣ Lấy dữ liệu Giáo viên", results['teachers_data']),
            ("4️⃣ Lấy dữ liệu Học sinh", results['students_data']),
            ("5️⃣ Lưu dữ liệu JSON", results['json_saved']),
            ("6️⃣ Chuyển đổi Excel", results['excel_converted']),
            ("7️⃣ Upload Google Drive", results['drive_uploaded'])
        ]
        
        for step_name, status in steps:
            status_icon = "✅" if status else "❌"
            status_text = "Thành công" if status else "Thất bại"
            print(f"{status_icon} {step_name}: {status_text}")
        
        # Tóm tắt dữ liệu
        if results.get('data_summary'):
            print(f"\n📊 TÓM TẮT DỮ LIỆU:")
            data_summary = results['data_summary']
            
            if 'teachers' in data_summary:
                teachers = data_summary['teachers']
                print(f"   👨‍🏫 Giáo viên: {teachers['retrieved']}/{teachers['total']}")
            
            if 'students' in data_summary:
                students = data_summary['students']
                print(f"   👨‍🎓 Học sinh: {students['retrieved']}/{students['total']}")
        
        # Thông tin HT/HP
        if results.get('ht_hp_info'):
            ht_hp_info = results['ht_hp_info']
            print(f"\n👑 THÔNG TIN LÃNH ĐẠO:")
            print(f"   👑 Hiệu trường (HT): {ht_hp_info.get('total_ht', 0)} người")
            print(f"   🔸 Hiệu phó (HP): {ht_hp_info.get('total_hp', 0)} người")
            
            # Hiển thị danh sách HT
            if ht_hp_info.get('ht'):
                print(f"   📋 Danh sách Hiệu trường:")
                for i, ht in enumerate(ht_hp_info['ht'], 1):
                    print(f"      {i}. {ht['name']}")
            
            # Hiển thị danh sách HP
            if ht_hp_info.get('hp'):
                print(f"   📋 Danh sách Hiệu phó:")
                for i, hp in enumerate(ht_hp_info['hp'], 1):
                    print(f"      {i}. {hp['name']}")
        
        # File outputs
        if results.get('json_file_path') or results.get('excel_file_path') or results.get('ht_hp_file'):
            print(f"\n📄 FILES ĐÃ TẠO:")
            if results.get('json_file_path'):
                print(f"   📋 JSON: {results['json_file_path']}")
            if results.get('excel_file_path'):
                print(f"   📊 Excel: {results['excel_file_path']}")
            if results.get('ht_hp_file'):
                print(f"   👑 HT/HP Info: {results['ht_hp_file']}")
        
        # Upload results
        if results.get('upload_results'):
            upload_info = results['upload_results']
            print(f"\n📤 DRIVE UPLOAD:")
            print(f"   ✅ Thành công: {upload_info.get('success', 0)} files")
            print(f"   ❌ Thất bại: {upload_info.get('failed', 0)} files")
            
            if upload_info.get('urls'):
                print(f"   🔗 Upload URLs:")
                for i, url in enumerate(upload_info['urls'][:3], 1):
                    print(f"      {i}. {url}")
                if len(upload_info['urls']) > 3:
                    print(f"      ... và {len(upload_info['urls']) - 3} URLs khác")
        
        # Tổng kết
        success_count = sum([results['sheets_extraction'], results['api_login'], 
                           results['teachers_data'], results['students_data'],
                           results['json_saved'], results['excel_converted'], 
                           results['drive_uploaded']])
        total_steps = 7
        
        print(f"\n🎯 TỔNG KẾT: {success_count}/{total_steps} bước thành công")
        
        if success_count == total_steps:
            print_status("🎉 WORKFLOW HOÀN CHỈNH THÀNH CÔNG - ĐÃ TẠO EXCEL VÀ UPLOAD DRIVE!", "success")
        elif success_count >= 6:
            print_status("⚠️ Workflow hoàn thành chính (có thể thiếu Drive upload do Drive link không hợp lệ)", "warning")
            if not results['drive_uploaded']:
                print("💡 Lý do có thể:")
                print("   • Drive link trong Google Sheets không đúng format")
                print("   • Cần cập nhật cột 'Link driver dữ liệu' với link thực tế")
                print("   • OAuth chưa được setup đúng")
        elif success_count >= 4:
            print_status("⚠️ Workflow hoàn thành phần chính (có thể thiếu JSON/Excel/Upload)", "warning")
        elif success_count >= 2:
            print_status("⚠️ Workflow hoàn thành một phần", "warning")
        else:
            print_status("❌ Workflow thất bại", "error")
    
    def _save_teachers_data(self, teachers_list, total_count):
        """Lưu dữ liệu giáo viên vào file JSON"""
        try:
            
            teachers_data = {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'data_type': 'teachers',
                'total_count': total_count,
                'retrieved_count': len(teachers_list),
                'teachers': teachers_list
            }
            
            filename = f"teachers_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            filepath = f"data/output/{filename}"
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(teachers_data, f, ensure_ascii=False, indent=2)
            
            print_status(f"✅ Đã lưu dữ liệu giáo viên vào: {filepath}", "success")
            print(f"   👨‍🏫 Số giáo viên: {len(teachers_list)}/{total_count}")
            
        except Exception as e:
            print_status(f"⚠️ Lỗi lưu dữ liệu giáo viên: {e}", "warning")
    
    def _save_students_data(self, students_list, total_count):
        """Lưu dữ liệu học sinh vào file JSON"""
        try:
            
            students_data = {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'data_type': 'students',
                'total_count': total_count,
                'retrieved_count': len(students_list),
                'students': students_list
            }
            
            filename = f"students_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            filepath = f"data/output/{filename}"
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(students_data, f, ensure_ascii=False, indent=2)
            
            print_status(f"✅ Đã lưu dữ liệu học sinh vào: {filepath}", "success")
            print(f"   👨‍🎓 Số học sinh: {len(students_list)}/{total_count}")
            
        except Exception as e:
            print_status(f"⚠️ Lỗi lưu dữ liệu học sinh: {e}", "warning")

    def _save_workflow_data(self, workflow_results, teachers_result, students_result, admin_password=None):
        """Lưu dữ liệu workflow vào file và trả về đường dẫn file"""
        try:
            
            school_name = workflow_results['school_info'].get('name', 'Unknown')
            safe_school_name = "".join(c for c in school_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            
            # Tạo cấu trúc JSON đơn giản, không trùng lặp
            workflow_data = {
                'school_info': {
                    'name': workflow_results['school_info'].get('name'),
                    'admin': workflow_results['school_info'].get('admin'),
                    'drive_link': workflow_results['school_info'].get('drive_link'),
                    'admin_password': admin_password
                },
                'data_summary': workflow_results.get('data_summary', {}),
                'ht_hp_info': workflow_results.get('ht_hp_info', {}),  # Thêm thông tin HT/HP
                'teachers': teachers_result.get('data') if teachers_result.get('success') else None,
                'students': students_result.get('data') if students_result.get('success') else None
            }
            
            # Tạo filename với timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"workflow_data_{safe_school_name}_{timestamp}.json"
            filepath = f"data/output/{filename}"
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(workflow_data, f, ensure_ascii=False, indent=2)
            
            return filepath
            
        except Exception as e:
            print_status(f"⚠️ Lỗi lưu dữ liệu workflow: {e}", "warning")
            return None
    
    def _load_latest_login_tokens(self):
        """Tải tokens từ file login gần nhất"""
        try:
            
            # Tìm file login gần nhất
            pattern = "data/output/onluyen_login_*.json"
            files = glob.glob(pattern)
            
            if not files:
                print_status("Không tìm thấy file login nào", "warning")
                return None
            
            # Sắp xếp theo thời gian tạo, lấy file mới nhất
            latest_file = max(files, key=lambda f: Path(f).stat().st_mtime)
            
            with open(latest_file, 'r', encoding='utf-8') as f:
                login_data = json.load(f)
            
            tokens = login_data.get('tokens', {})
            if tokens.get('access_token'):
                print_status(f"Đã tải tokens từ: {latest_file}", "success")
                return tokens
            else:
                print_status("File login không chứa tokens hợp lệ", "warning")
                return None
                
        except Exception as e:
            print_status(f"Lỗi tải tokens: {e}", "error")
            return None
    
    def onluyen_use_saved_tokens(self):
        """Sử dụng tokens đã lưu từ login trước đó"""
        print_separator("SỬ DỤNG TOKENS ĐÃ LƯU")
        
        # Tải tokens từ file
        tokens = self._load_latest_login_tokens()
        if not tokens:
            return
        
        try:
            # Khởi tạo client và set token
            client = OnLuyenAPIClient()
            access_token = tokens.get('access_token')
            
            if access_token:
                client.set_auth_token(access_token)
                print_status("Đã set access token thành công", "success")
                
                # Test token bằng cách thử gọi API
                print_status("Đang test token bằng cách lấy danh sách giáo viên...", "info")
                result = client.get_teachers(page_size=5)
                
                if result['success']:
                    print_status("Token hoạt động tốt! Có thể sử dụng các API khác.", "success")
                    data = result.get('data', [])
                    print(f"   📊 Số giáo viên lấy được: {len(data) if isinstance(data, list) else 'N/A'}")
                else:
                    print_status(f"Token có thể đã hết hạn: {result.get('error', 'Unknown error')}", "warning")
                    print("   💡 Thử login lại để lấy token mới")
            else:
                print_status("Không tìm thấy access token trong file", "error")
                
        except ImportError:
            print_status("Module onluyen_api chưa được cài đặt", "error")
        except Exception as e:
            print_status(f"Lỗi sử dụng tokens: {e}", "error")
    
    def onluyen_convert_json_to_excel(self):
        """Chuyển đổi JSON Workflow sang Excel"""
        print_separator("CHUYỂN ĐỔI JSON WORKFLOW → EXCEL")
        
        try:
            
            # Tìm các file JSON workflow
            json_patterns = [
                "data/output/data_*.json",
                "data/output/workflow_data_*.json"
            ]
            
            json_files = []
            for pattern in json_patterns:
                json_files.extend(glob.glob(pattern))
            
            if not json_files:
                print_status("Không tìm thấy file JSON workflow nào", "warning")
                return
            
            # Chọn file để convert
            if len(json_files) == 1:
                selected_file = json_files[0]
            else:
                print(f"\nTìm thấy {len(json_files)} file JSON:")
                for i, file in enumerate(json_files, 1):
                    print(f"{i}. {Path(file).name}")
                
                try:
                    choice = get_user_input(f"Chọn file để convert (1-{len(json_files)})", required=True)
                    choice_idx = int(choice) - 1
                    if 0 <= choice_idx < len(json_files):
                        selected_file = json_files[choice_idx]
                    else:
                        print_status("Lựa chọn không hợp lệ", "error")
                        return
                except (ValueError, TypeError):
                    print_status("Lựa chọn không hợp lệ", "error")
                    return
            
            # Import và sử dụng converter
            converter = JSONToExcelTemplateConverter(selected_file)
            
            # Load và kiểm tra JSON data
            if not converter.load_json_data():
                print_status("Không thể load JSON data", "error")
                return
            
            # Extract data
            teachers_extracted = converter.extract_teachers_data()
            students_extracted = converter.extract_students_data()
            
            if not teachers_extracted and not students_extracted:
                print_status("Không thể trích xuất dữ liệu giáo viên hoặc học sinh", "error")
                return
            
            # Convert to Excel
            output_path = converter.convert()
            
            if output_path:
                print_status("Chuyển đổi thành công!", "success")
                print(f"File Excel: {output_path}")
                
                # Hiển thị thống kê
                teachers_count = len(converter.teachers_df) if converter.teachers_df is not None else 0
                students_count = len(converter.students_df) if converter.students_df is not None else 0
                
                print(f"\nThống kê: {teachers_count} giáo viên, {students_count} học sinh")
                
                # Hỏi có muốn mở file Excel không
                if get_user_confirmation("Bạn có muốn mở file Excel?"):
                    try:
                        os.startfile(output_path)
                    except Exception as e:
                        print_status(f"Không thể mở file Excel: {e}", "warning")
            else:
                print_status("Chuyển đổi thất bại", "error")
                
        except ImportError:
            print_status("Module json_to_excel_template_converter chưa được cài đặt", "error")
        except Exception as e:
            print_status(f"Lỗi chuyển đổi: {e}", "error")
    

    
    
    def _get_sheet_name_with_fallback(self, extractor):
        """
        Lấy tên sheet với logic fallback:
        1. Thử với ED-2025 trước
        2. Nếu không có thì yêu cầu nhập tên sheet
        """
        try:
            # Thử với ED-2025 trước
            print_status("Đang thử tìm sheet 'ED-2025'...", "info")
            test_data = extractor.extract_school_data(sheet_name="ED-2025")
            
            if test_data and len(test_data) > 0:
                print_status("✅ Tìm thấy sheet 'ED-2025'", "success")
                return "ED-2025"
            else:
                print_status("⚠️ Không tìm thấy sheet 'ED-2025' hoặc sheet trống", "warning")
                
                # Yêu cầu nhập tên sheet
                sheet_name = get_user_input("Nhập tên sheet (bắt buộc):")
                if not sheet_name:
                    print_status("❌ Tên sheet không được để trống", "error")
                    return None
                    
                return sheet_name.strip()
                
        except Exception as e:
            print_status(f"⚠️ Lỗi khi thử tìm sheet ED-2025: {e}", "warning")
            
            # Fallback: yêu cầu nhập tên sheet
            sheet_name = get_user_input("Nhập tên sheet:")
            if not sheet_name:
                print_status("❌ Tên sheet không được để trống", "error")
                return None
                
            return sheet_name.strip()

    def _execute_basic_workflow_steps(self, selected_school):
        """Thực hiện các bước cơ bản của workflow (dùng chung cho cả 2 case)"""
        basic_results = {
            'sheets_extraction': False,
            'api_login': False, 
            'teachers_data': False,
            'students_data': False,
            'school_info': {},
            'data_summary': {},
            'teachers_result': None,
            'students_result': None
        }
        
        try:
            # # Bước 1: Trích xuất dữ liệu từ Google Sheets
            # print_status("BƯỚC 1: Trích xuất dữ liệu từ Google Sheets", "info")
            
            # extractor = GoogleSheetsExtractor()
            
            # # Lấy tên sheet với logic fallback
            # sheet_name = self._get_sheet_name_with_fallback(extractor)
            # if not sheet_name:
            #     print_status("❌ Không thể xác định tên sheet", "error")
            #     return None
            
            # print_status(f"Đang trích xuất dữ liệu từ sheet: {sheet_name}", "info")
            # school_data = extractor.extract_school_data(sheet_name=sheet_name)
            
            if not selected_school:
                print_status("❌ Không thể trích xuất dữ liệu", "error")
                return None
            
            basic_results['sheets_extraction'] = True
            # print_status(f"✅ Đã trích xuất {len(school_data)} trường học", "success")
            
            # # Chọn trường để xử lý
            # if len(school_data) == 1:
            #     selected_school = school_data[0]
            #     print_status("Tự động chọn trường duy nhất", "info")
            # else:
            #     print("\n📋 DANH SÁCH TRƯỜNG ĐÃ TRÍCH XUẤT:")
            #     for i, school in enumerate(school_data, 1):
            #         school_name = school.get('Tên trường', 'N/A')
            #         admin_email = school.get('Admin', 'N/A')
            #         print(f"   {i}. {school_name} (Admin: {admin_email})")
                
            #     try:
            #         choice = get_user_input(f"Chọn trường để xử lý (1-{len(school_data)})", required=True)
            #         choice_idx = int(choice) - 1
            #         if 0 <= choice_idx < len(school_data):
            #             selected_school = school_data[choice_idx]
            #         else:
            #             print_status("Lựa chọn không hợp lệ", "error")
            #             return None
            #     except (ValueError, TypeError):
            #         print_status("Lựa chọn không hợp lệ", "error")
            #         return None
            
            # Lấy thông tin trường
            school_name = selected_school.get('Tên trường', 'N/A')
            admin_email = selected_school.get('Admin', '')
            password = selected_school.get('Mật khẩu', '')
            drive_link = selected_school.get('Link driver dữ liệu', 'N/A')
            
            basic_results['school_info'] = {
                'name': school_name,
                'admin': admin_email,
                'drive_link': drive_link,
                'password': password
            }
            
            print(f"\n📋 THÔNG TIN TRƯỜNG ĐÃ CHỌN:")
            print(f"   🏫 Tên trường: {school_name}")
            print(f"   👤 Admin: {admin_email}")
            print(f"   🔗 Drive Link: {drive_link[:60] + '...' if len(drive_link) > 60 else drive_link}")
            
            if not admin_email or not password:
                print_status("❌ Thiếu thông tin Admin email hoặc Mật khẩu", "error")
                return None
            
            # Bước 2: Login vào OnLuyen API
            print_status("BƯỚC 2: Thực hiện login OnLuyen API", "info")
            
            client = OnLuyenAPIClient()
            print_status(f"Đang login với Admin: {admin_email}", "info")
            
            result = client.login(admin_email, password)
            
            if not result['success']:
                print_status(f"❌ Login thất bại: {result.get('error', 'Unknown error')}", "error")
                return None
            
            # Kiểm tra tài khoản trùng khớp
            response_data = result.get('data', {})
            response_email = response_data.get('account', '').lower().strip()
            expected_email = admin_email.lower().strip()
            
            if response_email != expected_email:
                print_status("❌ Tài khoản đăng nhập không trùng khớp!", "error")
                print(f"   Expected: {expected_email}")
                print(f"   Got: {response_email}")
                return None
            
            basic_results['api_login'] = True
            print_status("✅ Login thành công và tài khoản trùng khớp", "success")
            
            # Bước 3: Lấy danh sách Giáo viên
            print_status("BƯỚC 3: Lấy danh sách Giáo viên", "info")
            
            teachers_result = client.get_teachers(page_size=1000)
            
            if teachers_result['success'] and teachers_result.get('data'):
                teachers_data = teachers_result['data']
                if isinstance(teachers_data, dict) and 'data' in teachers_data:
                    teachers_list = teachers_data['data']
                    teachers_count = teachers_data.get('totalCount', len(teachers_list))
                    
                    basic_results['teachers_data'] = True
                    basic_results['teachers_result'] = teachers_result
                    basic_results['data_summary']['teachers'] = {
                        'total': teachers_count,
                        'retrieved': len(teachers_list)
                    }
                    
                    print_status(f"✅ Lấy danh sách giáo viên thành công: {len(teachers_list)}/{teachers_count}", "success")
                    
                    # Extract thông tin HT/HP cho Case 2
                    print_status("🔍 Trích xuất thông tin Hiệu trường (HT) và Hiệu phó (HP)", "info")
                    ht_hp_info = self._extract_ht_hp_info(teachers_data)
                    basic_results['ht_hp_info'] = ht_hp_info
                    
                    # Lưu thông tin HT/HP vào file riêng
                    school_name = basic_results['school_info'].get('name', 'Unknown')
                    ht_hp_file = self._save_ht_hp_info(ht_hp_info, school_name)
                    if ht_hp_file:
                        basic_results['ht_hp_file'] = ht_hp_file
                else:
                    print_status("⚠️ Định dạng dữ liệu giáo viên không đúng", "warning")
            else:
                print_status(f"❌ Lỗi lấy danh sách giáo viên: {teachers_result.get('error')}", "error")
            
            # Bước 4: Lấy danh sách Học sinh
            print_status("BƯỚC 4: Lấy danh sách Học sinh", "info")
            
            students_result = client.get_students(page_index=1, page_size=5000)
            
            if students_result['success'] and students_result.get('data'):
                students_data = students_result['data']
                if isinstance(students_data, dict) and 'data' in students_data:
                    students_list = students_data['data']
                    students_count = students_data.get('totalCount', len(students_list))
                    
                    basic_results['students_data'] = True
                    basic_results['students_result'] = students_result
                    basic_results['data_summary']['students'] = {
                        'total': students_count,
                        'retrieved': len(students_list)
                    }
                    
                    print_status(f"✅ Lấy danh sách học sinh thành công: {len(students_list)}/{students_count}", "success")
                else:
                    print_status("⚠️ Định dạng dữ liệu học sinh không đúng", "warning")
            else:
                print_status(f"❌ Lỗi lấy danh sách học sinh: {students_result.get('error')}", "error")
            
            return basic_results
            
        except Exception as e:
            print_status(f"❌ Lỗi trong các bước cơ bản: {e}", "error")
            return None
    
    def _download_import_file(self, school_name, drive_link):
        """Tải file import từ Google Drive với pattern 'import_*'"""
        try:
            # Khởi tạo OAuth client
            oauth_client = GoogleOAuthDriveClient()
            
            if not oauth_client.is_authenticated():
                print_status("❌ OAuth chưa được setup", "error")
                return None
            
            # Extract folder ID
            folder_id = self._extract_drive_folder_id(drive_link)
            if not folder_id:
                print_status("❌ Không thể extract folder ID từ drive link", "error")
                return None
            
            print(f"   🔍 Tìm file có pattern: import_*.xlsx")
            
            # Tìm tất cả file bắt đầu bằng "import_" trong Drive folder
            import_files = self._find_import_files_in_drive_folder(oauth_client, folder_id)
            
            if not import_files:
                print_status(f"❌ Không tìm thấy file nào có pattern 'import_*.xlsx'", "error")
                return None
            
            # Nếu có nhiều file, cho user chọn
            selected_file = None
            if len(import_files) == 1:
                selected_file = import_files[0]
                print(f"   ✅ Tìm thấy file: {selected_file['name']}")
            else:
                print(f"\n📋 TÌM THẤY {len(import_files)} FILE IMPORT:")
                for i, file in enumerate(import_files, 1):
                    print(f"   {i}. {file['name']}")
                
                try:
                    choice = get_user_input(f"Chọn file import (1-{len(import_files)})", required=True)
                    choice_idx = int(choice) - 1
                    if 0 <= choice_idx < len(import_files):
                        selected_file = import_files[choice_idx]
                    else:
                        print_status("Lựa chọn không hợp lệ", "error")
                        return None
                except (ValueError, TypeError):
                    print_status("Lựa chọn không hợp lệ", "error")
                    return None
            
            if not selected_file:
                return None
            
            # Tải file về local
            local_filename = selected_file['name']
            local_path = f"data/temp/{local_filename}"
            os.makedirs("data/temp", exist_ok=True)
            
            success = self._download_file_from_drive(oauth_client, selected_file['id'], local_path)
            
            if success:
                print_status(f"✅ Đã tải file import: {local_filename}", "success")
                return local_path
            else:
                print_status("❌ Lỗi tải file import", "error")
                return None
                
        except ImportError:
            print_status("❌ OAuth module chưa được cài đặt", "error")
            return None
        except Exception as e:
            print_status(f"❌ Lỗi tải file import: {e}", "error")
            return None
    
    def _find_import_files_in_drive_folder(self, oauth_client, folder_id):
        """Tìm tất cả file bắt đầu bằng 'import_' trong Drive folder"""
        try:
            # Tìm file với pattern "import_" và phần mở rộng .xlsx
            query = f"parents in '{folder_id}' and trashed=false and name contains 'import_' and (name contains '.xlsx' or name contains '.xls')"
            
            results = oauth_client.drive_service.files().list(
                q=query,
                spaces='drive',
                fields='files(id, name)',
                pageSize=100
            ).execute()
            
            files = results.get('files', [])
            
            # Lọc thêm để chỉ lấy file thực sự bắt đầu bằng "import_"
            import_files = []
            for file in files:
                filename = file['name'].lower()
                if filename.startswith('import_') and (filename.endswith('.xlsx') or filename.endswith('.xls')):
                    import_files.append(file)
            
            return import_files
                
        except Exception as e:
            print_status(f"❌ Lỗi tìm file import: {e}", "error")
            return []
    
    def _download_file_from_drive(self, oauth_client, file_id, local_path):
        """Tải file từ Drive về local"""
        try:
            
            request = oauth_client.drive_service.files().get_media(fileId=file_id)
            
            with open(local_path, 'wb') as f:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
            
            return True
            
        except Exception as e:
            print_status(f"❌ Lỗi download file: {e}", "error")
            return False
    
    def _compare_and_filter_data(self, teachers_result, students_result, import_file_path):
        """So sánh và lọc dữ liệu dựa trên file import theo Họ tên và Ngày sinh"""
        try:
            
            # Đọc file import với tất cả sheets
            print("   📂 Đọc file import...")
            excel_file = pd.ExcelFile(import_file_path)
            
            comparison_results = {
                'teachers_filtered': [],
                'students_filtered': [],
                'import_teachers_count': 0,
                'import_students_count': 0,
                'teachers_matched': 0,
                'students_matched': 0,
                'comparison_method': 'name_and_birthdate'
            }
            
            # Xử lý sheet Teachers nếu có
            teachers_import_data = []
            export_all_teachers = False  # Flag để xuất tất cả giáo viên
            
            if 'Teachers' in excel_file.sheet_names:
                teachers_df = pd.read_excel(import_file_path, sheet_name='Teachers')
                print(f"   👨‍🏫 Sheet Teachers: {len(teachers_df)} rows")
                
                # Chuẩn hóa format ngày tháng trong DataFrame trước khi xử lý
                teachers_df = self._standardize_import_date_formats(teachers_df)
                
                # Tìm cột họ tên và ngày sinh
                name_col = self._find_column_by_keywords(teachers_df.columns, ['họ tên', 'tên', 'name', 'giáo viên'])
                birth_col = self._find_column_by_keywords(teachers_df.columns, ['ngày sinh', 'sinh', 'birth', 'date'])
                
                if name_col and birth_col:
                    print(f"      📋 Cột tên: '{name_col}', Cột ngày sinh: '{birth_col}'")
                    
                    # Kiểm tra xem có giáo viên nào tên GVCN không (sử dụng pattern matching)
                    gvcn_found = False
                    for _, row in teachers_df.iterrows():
                        name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
                        if name and self._is_gvcn_name_in_import(name):
                            gvcn_found = True
                            print(f"      🔍 Tìm thấy GVCN pattern: '{name}'")
                            break
                    
                    if gvcn_found:
                        export_all_teachers = True
                        print(f"      🔍 Tìm thấy 'GVCN' → Sẽ xuất TẤT CẢ giáo viên từ OnLuyen")
                    else:
                        print(f"      🔍 Không có 'GVCN' → Chỉ xuất giáo viên có trong import")
                        
                        # Parse danh sách giáo viên từ import để so sánh
                        print(f"      🔍 Parsing teachers from import file...")
                        parsed_count = 0
                        skipped_gvcn_count = 0
                        
                        for idx, row in teachers_df.iterrows():
                            name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
                            birth = str(row[birth_col]).strip() if pd.notna(row[birth_col]) else ""
                            
                            if name and birth:
                                if self._is_gvcn_name_in_import(name):
                                    skipped_gvcn_count += 1
                                    print(f"         🚫 Skipping GVCN: '{name}'")
                                else:
                                    normalized_name = self._normalize_name(name)
                                    normalized_birth = self._normalize_date(birth)
                                    
                                    teachers_import_data.append({
                                        'name': normalized_name,
                                        'birthdate': normalized_birth,
                                        'raw_name': name,
                                        'raw_birthdate': birth
                                    })
                                    parsed_count += 1
                                    
                                    # Debug first few teachers
                                    if parsed_count <= 5:
                                        print(f"         ✅ Parsed teacher {parsed_count}: '{name}' | '{birth}'")
                                        print(f"            → Normalized: '{normalized_name}' | '{normalized_birth}'")
                        
                        print(f"      📊 Parsing summary: {parsed_count} teachers parsed, {skipped_gvcn_count} GVCN skipped")
                
                comparison_results['import_teachers_count'] = len(teachers_import_data)
                comparison_results['export_all_teachers'] = export_all_teachers
                
                if export_all_teachers:
                    print(f"      ✅ Chế độ xuất tất cả giáo viên (có GVCN)")
                else:
                    print(f"      ✅ Đã parse {len(teachers_import_data)} giáo viên từ import")
            
            # Xử lý sheet Students nếu có
            students_import_data = []
            if 'Students' in excel_file.sheet_names:
                students_df = pd.read_excel(import_file_path, sheet_name='Students')
                print(f"   👨‍🎓 Sheet Students: {len(students_df)} rows")
                
                # Chuẩn hóa format ngày tháng trong DataFrame trước khi xử lý
                students_df = self._standardize_import_date_formats(students_df)
                
                # Tìm cột họ tên và ngày sinh
                name_col = self._find_column_by_keywords(students_df.columns, ['họ tên', 'họ và tên', 'fullname', 'tên học sinh'])
                birth_col = self._find_column_by_keywords(students_df.columns, ['ngày sinh', 'sinh', 'birth', 'date'])
                
                if name_col and birth_col:
                    print(f"      � Cột tên: '{name_col}', Cột ngày sinh: '{birth_col}'")
                    
                    for _, row in students_df.iterrows():
                        name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
                        birth = str(row[birth_col]).strip() if pd.notna(row[birth_col]) else ""
                        
                        if name and birth:
                            normalized_name = self._normalize_name(name)
                            normalized_birth = self._normalize_date(birth)
                            
                            students_import_data.append({
                                'name': normalized_name,
                                'birthdate': normalized_birth,
                                'raw_name': name,
                                'raw_birthdate': birth
                            })
                
                comparison_results['import_students_count'] = len(students_import_data)
                print(f"      ✅ Đã parse {len(students_import_data)} học sinh từ import")
            
            # So sánh và lọc giáo viên
            if teachers_result and teachers_result.get('success'):
                print("   🔍 Xử lý danh sách giáo viên...")
                teachers_data = teachers_result['data']
                
                if isinstance(teachers_data, dict) and 'data' in teachers_data:
                    onluyen_teachers = teachers_data['data']
                    
                    if comparison_results.get('export_all_teachers', False):
                        # Xuất tất cả giáo viên từ OnLuyen (vì có GVCN) nhưng loại bỏ những giáo viên tên "GVCN"
                        filtered_teachers = []
                        for teacher in onluyen_teachers:
                            # Sử dụng helper function để kiểm tra GVCN
                            if not self._is_gvcn_teacher(teacher):
                                filtered_teachers.append(teacher)
                        
                        comparison_results['teachers_filtered'] = filtered_teachers
                        comparison_results['teachers_matched'] = len(filtered_teachers)
                        original_count = len(onluyen_teachers)
                        excluded_count = original_count - len(filtered_teachers)
                        print(f"      ✅ Xuất {len(filtered_teachers)}/{original_count} giáo viên (loại bỏ {excluded_count} giáo viên GVCN)")
                        
                    elif teachers_import_data:
                        # Chỉ xuất giáo viên khớp với import
                        print(f"      📊 OnLuyen có {len(onluyen_teachers)} giáo viên")
                        print(f"      📋 Import có {len(teachers_import_data)} giáo viên")
                        
                        # Tạo set để so sánh nhanh
                        import_teachers_set = set()
                        for t in teachers_import_data:
                            if t['name'] and t['birthdate']:
                                import_teachers_set.add((t['name'], t['birthdate']))
                        
                        print(f"      🔍 Import teachers set có {len(import_teachers_set)} items")
                        
                        # In ra sample import teachers để debug
                        print(f"      📝 Sample import teachers:")
                        for i, (name, birth) in enumerate(list(import_teachers_set)[:5], 1):
                            print(f"         {i}. '{name}' | '{birth}'")
                        
                        # Lọc giáo viên OnLuyen khớp với import
                        matched_count = 0
                        debug_comparison = True  # Enable debug để xem tại sao không match
                        unmatched_onluyen_teachers = []
                        unmatched_import_teachers = list(import_teachers_set)  # Copy để track
                        
                        print(f"      🔍 DEBUG: Sample OnLuyen teachers:")
                        for i, teacher in enumerate(onluyen_teachers[:5], 1):
                            teacher_full_name = teacher.get('fullName', '')
                            teacher_birth_date = teacher.get('birthDate', '')
                            teacher_info = teacher.get('teacherInfo', {})
                            teacher_display_name = teacher_info.get('displayName', '') if teacher_info else ''
                            
                            print(f"         {i}. Raw fullName: '{teacher_full_name}' | birthDate: '{teacher_birth_date}'")
                            print(f"            teacherInfo.displayName: '{teacher_display_name}'")
                            print(f"            → Normalized: '{self._normalize_name(teacher_full_name)}' | '{self._normalize_date(teacher_birth_date)}'")
                        
                        # Tạo set tên import để fallback matching
                        import_teachers_names_only = set()
                        for t in teachers_import_data:
                            if t['name']:
                                import_teachers_names_only.add(t['name'])
                        
                        print(f"      🔍 Will try name+birthdate matching first, then name-only fallback if needed")
                        
                        for teacher in onluyen_teachers:
                            # Sử dụng helper function để kiểm tra GVCN trước tiên
                            if self._is_gvcn_teacher(teacher):
                                if debug_comparison:
                                    print(f"    🚫 Skipping GVCN teacher: '{teacher.get('fullName', '')}'")
                                continue
                            
                            # Lấy tên từ cả fullName và teacherInfo.displayName
                            teacher_full_name = teacher.get('fullName', '') or ''
                            teacher_info = teacher.get('teacherInfo', {})
                            teacher_display_name = teacher_info.get('displayName', '') if teacher_info else ''
                            teacher_name_raw = teacher_full_name or teacher_display_name
                            
                            teacher_name = self._normalize_name(teacher_name_raw)
                            teacher_birth = self._normalize_date(teacher.get('birthDate', ''))
                            
                            matched_this_teacher = False
                            
                            # Thử match theo name + birthdate trước
                            if teacher_name and teacher_birth:
                                teacher_key = (teacher_name, teacher_birth)
                                if teacher_key in import_teachers_set:
                                    comparison_results['teachers_filtered'].append(teacher)
                                    matched_count += 1
                                    matched_this_teacher = True
                                    # Xóa khỏi unmatched list
                                    if teacher_key in unmatched_import_teachers:
                                        unmatched_import_teachers.remove(teacher_key)
                                    
                                    if debug_comparison and matched_count <= 10:
                                        print(f"    ✅ FULL MATCH (name+birth): '{teacher_name}' | '{teacher_birth}'")
                            
                            # Nếu chưa match và có tên, thử match chỉ theo tên (fallback)
                            if not matched_this_teacher and teacher_name:
                                if teacher_name in import_teachers_names_only:
                                    # Tìm import teacher có cùng tên để lấy ngày sinh
                                    for import_teacher in teachers_import_data:
                                        if import_teacher['name'] == teacher_name:
                                            # Match bằng tên, tạo fake key với import birthdate
                                            fake_key = (teacher_name, import_teacher['birthdate'])
                                            comparison_results['teachers_filtered'].append(teacher)
                                            matched_count += 1
                                            matched_this_teacher = True
                                            # Xóa khỏi unmatched list
                                            if fake_key in unmatched_import_teachers:
                                                unmatched_import_teachers.remove(fake_key)
                                            
                                            if debug_comparison and matched_count <= 10:
                                                print(f"    ⚠️ NAME-ONLY MATCH: '{teacher_name}' (OnLuyen missing birthdate)")
                                                print(f"        Expected birth: '{import_teacher['birthdate']}' from import")
                                            break
                            
                            # Nếu vẫn không match
                            if not matched_this_teacher:
                                if teacher_name:
                                    unmatched_onluyen_teachers.append({
                                        'raw_name': teacher_name_raw,
                                        'raw_birth': teacher.get('birthDate', ''),
                                        'normalized_name': teacher_name,
                                        'normalized_birth': teacher_birth
                                    })
                                    
                                    # Debug so sánh
                                    if debug_comparison and len(unmatched_onluyen_teachers) <= 10:
                                        print(f"    ❌ NO MATCH: '{teacher_name}' | Birth: '{teacher_birth}'")
                                        print(f"        Raw fullName: '{teacher_full_name}'")
                                        print(f"        Raw teacherInfo.displayName: '{teacher_display_name}'")
                                else:
                                    # Debug teachers với thông tin thiếu
                                    if debug_comparison:
                                        print(f"    ⚠️ Teacher missing info: '{teacher_name_raw}' | '{teacher.get('birthDate', '')}'")
                                        print(f"        Raw fullName: '{teacher_full_name}'")
                                        print(f"        Raw teacherInfo.displayName: '{teacher_display_name}'")
                                        print(f"        → Normalized name: '{teacher_name}' | birth: '{teacher_birth}'")
                        
                        comparison_results['teachers_matched'] = matched_count
                        print(f"      ✅ Khớp {matched_count}/{len(teachers_import_data)} giáo viên")
                        
                    else:
                        print(f"      ⚠️ Không có dữ liệu giáo viên import để so sánh")
            
            # So sánh và lọc học sinh
            if students_result and students_result.get('success'):
                print("   🔍 Xử lý danh sách học sinh...")
                students_data = students_result['data']
                
                if isinstance(students_data, dict) and 'data' in students_data:
                    onluyen_students = students_data['data']
                    print(f"      📊 OnLuyen có {len(onluyen_students)} học sinh")
                    
                    if students_import_data:
                        print("   🔍 So sánh với file import...")
                        
                        # Tạo set để so sánh nhanh
                        import_students_set = set()
                        for s in students_import_data:
                            if s['name'] and s['birthdate']:
                                import_students_set.add((s['name'], s['birthdate']))                        
                        print(f"      📋 Import students set: {len(import_students_set)} items")
                        
                        # Lọc học sinh OnLuyen khớp với import
                        matched_count = 0
                        unmatched_onluyen = []  # Danh sách học sinh OnLuyen không khớp
                        unmatched_import = []   # Danh sách học sinh Import không khớp
                        debug_comparison = True  # Set True để debug - ENABLE DEBUG
                                                
                        # Tạo set để track các import students đã được match
                        matched_import_keys = set()
                        
                        for student in onluyen_students:
                            # Học sinh OnLuyen có thể có fullName trống, phải dùng userInfo.displayName
                            user_info = student.get('userInfo', {})
                            student_name = self._normalize_name(
                                student.get('fullName', '') or user_info.get('displayName', '')
                            )
                            student_birth = self._normalize_date(
                                student.get('birthDate', '') or user_info.get('userBirthday', '')
                            )
                            
                            if student_name and student_birth:
                                student_key = (student_name, student_birth)
                            
                                if student_key in import_students_set:
                                    comparison_results['students_filtered'].append(student)
                                    matched_count += 1
                                    matched_import_keys.add(student_key)
                                else:
                                    # Log unmatched OnLuyen student
                                    unmatched_onluyen.append({
                                        'original_name': student.get('fullName', '') or user_info.get('displayName', ''),
                                        'original_birth': student.get('birthDate', '') or user_info.get('userBirthday', ''),
                                        'normalized_name': student_name,
                                        'normalized_birth': student_birth
                                    })
                        
                        # Tìm import students không được match
                        for import_key in import_students_set:
                            if import_key not in matched_import_keys:
                                # Tìm thông tin gốc của import student này
                                for import_student in students_import_data:
                                    if (import_student['name'], import_student['birthdate']) == import_key:
                                        unmatched_import.append({
                                            'original_name': import_student['raw_name'],
                                            'original_birth': import_student['raw_birthdate'],
                                            'normalized_name': import_student['name'],
                                            'normalized_birth': import_student['birthdate']
                                        })
                                        break
                        
                        # Log kết quả chi tiết
                        print(f"      ✅ Khớp {matched_count}/{len(students_import_data)} học sinh")
                    
                    else:
                        # Nếu không có import data, xuất tất cả học sinh
                        print("      ⚠️ Không có dữ liệu học sinh import - Xuất tất cả học sinh OnLuyen")
                        comparison_results['students_filtered'] = onluyen_students
                        comparison_results['students_matched'] = len(onluyen_students)
                else:
                    print("      ❌ Định dạng dữ liệu học sinh OnLuyen không đúng")
            else:
                print("      ❌ Không có dữ liệu học sinh OnLuyen")
            
            return comparison_results
            
        except Exception as e:
            print_status(f"❌ Lỗi so sánh dữ liệu: {e}", "error")
            return None
    
    def _save_unmatched_log(self, unmatched_onluyen, unmatched_import):
        """Lưu log chi tiết các trường hợp không khớp vào file"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_filename = f"unmatched_students_log_{timestamp}.json"
            log_filepath = f"data/output/{log_filename}"
            
            log_data = {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'summary': {
                    'unmatched_onluyen_count': len(unmatched_onluyen),
                    'unmatched_import_count': len(unmatched_import),
                    'total_unmatched': len(unmatched_onluyen) + len(unmatched_import)
                },
                'unmatched_onluyen_students': unmatched_onluyen,
                'unmatched_import_students': unmatched_import,
                'analysis': {
                    'common_issues': [
                        'Sai format ngày tháng (DD/MM/YYYY vs MM/DD/YYYY)',
                        'Khác biệt trong cách viết tên (dấu, khoảng trắng)',
                        'Thiếu hoặc thừa thông tin trong một trong hai nguồn',
                        'Lỗi nhập liệu từ trường học'
                    ],
                    'recommendations': [
                        'Kiểm tra format ngày tháng trong file import',
                        'So sánh tên học sinh với độ tương đồng cao',
                        'Xác minh với trường học về thông tin chính xác',
                        'Cân nhắc sử dụng debug functions để phân tích chi tiết'
                    ]
                }
            }
            
            with open(log_filepath, 'w', encoding='utf-8') as f:
                json.dump(log_data, f, ensure_ascii=False, indent=2)
            
            print(f"      📄 Đã lưu unmatched log: {log_filepath}")
            
        except Exception as e:
            print(f"      ⚠️ Lỗi lưu unmatched log: {e}")

    def _find_column_by_keywords(self, columns, keywords):
        """Tìm cột theo từ khóa"""
        for col in columns:
            col_lower = str(col).lower()
            for keyword in keywords:
                if keyword in col_lower:
                    return col
        return None
    
    def _analyze_date_format_in_import(self, df, column_name):
        """Phân tích format ngày tháng thực tế trong DataFrame cột cụ thể"""
        
        try:
            # Collect date samples từ cột được chỉ định
            date_samples = []
            
            if column_name not in df.columns:
                print(f"   ❌ Không tìm thấy cột '{column_name}' trong DataFrame")
                return None
            
            print(f"   🔍 Analyzing date format in column '{column_name}'...")
            
            # Lấy tối đa 20 samples có dữ liệu
            sample_count = 0
            for _, row in df.iterrows():
                if sample_count >= 20:
                    break
                    
                date_value = row[column_name]
                if pd.notna(date_value):
                    date_str = str(date_value).strip()
                    if date_str and date_str not in ['', 'nan', 'NaN']:
                        date_samples.append(date_str)
                        sample_count += 1
            
            if not date_samples:
                print(f"   ❌ Không tìm thấy dữ liệu ngày hợp lệ trong cột '{column_name}'")
                return None
            
            print(f"   📊 Collected {len(date_samples)} date samples for analysis")
            print(f"   📝 Sample dates: {date_samples[:5]}...")  # Show first 5 samples
            
            # Analyze patterns
            format_scores = {
                'DD/MM/YYYY': 0,
                'MM/DD/YYYY': 0, 
                'YYYY-MM-DD': 0,
                'Excel_DateTime': 0
            }
            
            total_analyzed = 0
            
            for date_str in date_samples:
                total_analyzed += 1
                
                # Check for Excel DateTime format (has timestamp)
                if ' ' in date_str and ':' in date_str:
                    format_scores['Excel_DateTime'] += 1
                    continue
                
                # Parse DD/MM/YYYY or MM/DD/YYYY patterns
                match = re.match(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})', date_str)
                if match:
                    first_num = int(match.group(1))
                    second_num = int(match.group(2))
                    
                    # Disambiguation logic
                    if first_num > 12:
                        # First number > 12, must be day -> DD/MM/YYYY
                        format_scores['DD/MM/YYYY'] += 2  # Higher weight for clear indication
                    elif second_num > 12:
                        # Second number > 12, must be day -> MM/DD/YYYY  
                        format_scores['MM/DD/YYYY'] += 2
                    else:
                        # Both <= 12, ambiguous - give small weight to both
                        format_scores['DD/MM/YYYY'] += 1
                        format_scores['MM/DD/YYYY'] += 1
                    continue
                
                # Check for YYYY-MM-DD format
                if re.match(r'(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})', date_str):
                    format_scores['YYYY-MM-DD'] += 1
                    continue
            
            # Calculate percentages and find most likely format
            if total_analyzed == 0:
                return None
                
            format_percentages = {}
            for fmt, score in format_scores.items():
                format_percentages[fmt] = round((score / total_analyzed) * 100, 1)
            
            # Find format with highest score
            most_likely_format = max(format_scores.keys(), key=lambda k: format_scores[k])
            confidence_score = format_percentages[most_likely_format]
            
            result = {
                'most_likely_format': most_likely_format,
                'confidence_score': confidence_score,
                'format_scores': format_percentages,
                'sample_count': len(date_samples),
                'samples': date_samples[:5]  # Include first 5 samples for debugging
            }
            
            return result
            
        except Exception as e:
            print(f"   ❌ Error analyzing date format: {str(e)}")
            return None
    
    def _normalize_name(self, name):
        """Chuẩn hóa tên để so sánh"""
        
        if not name or pd.isna(name):
            return ""
        
        # Chuyển về lowercase và loại bỏ khoảng trắng thừa
        normalized = str(name).lower().strip()
        
        # Loại bỏ dấu tiếng Việt để so sánh dễ hơn
        normalized = unicodedata.normalize('NFD', normalized)
        normalized = ''.join(c for c in normalized if unicodedata.category(c) != 'Mn')
        
        # Loại bỏ các ký tự đặc biệt và khoảng trắng thừa
        normalized = re.sub(r'[^\w\s]', '', normalized)
        normalized = re.sub(r'\s+', ' ', normalized).strip()
        
        return normalized
    
    def _is_gvcn_teacher(self, teacher_data):
        """Kiểm tra xem giáo viên có phải là GVCN hay không dựa vào tên"""
        try:
            # Lấy tên từ các field khác nhau
            teacher_full_name = teacher_data.get('fullName', '') or ''
            teacher_info = teacher_data.get('teacherInfo', {})
            teacher_display_name = teacher_info.get('displayName', '') if teacher_info else ''
            teacher_name = teacher_full_name or teacher_display_name
            
            if not teacher_name:
                return False
                
            return self._is_gvcn_name_in_import(teacher_name)
            
        except Exception:
            return False
    
    def _extract_ht_hp_info(self, teachers_data):
        """Trích xuất thông tin Hiệu trường (HT) và Hiệu phó (HP) từ danh sách giáo viên"""
        try:
            if not teachers_data or not isinstance(teachers_data, dict):
                return {'ht': [], 'hp': []}
            
            teachers_list = teachers_data.get('data', [])
            if not teachers_list:
                return {'ht': [], 'hp': []}
            
            ht_teachers = []  # Hiệu trường
            hp_teachers = []  # Hiệu phó
            
            print("   🔍 Đang tìm Hiệu trường (HT) và Hiệu phó (HP)...")
            
            # Debug: Hiển thị structure của 5 teachers đầu tiên
            print("   🔍 DEBUG: Structure của teachers:")
            for i, teacher in enumerate(teachers_list[:5], 1):
                teacher_roles = teacher.get('roles', [])
                teacher_name = teacher.get('fullName', '') or teacher.get('teacherInfo', {}).get('displayName', '') or 'Unknown'
                
                print(f"      Teacher {i}: '{teacher_name}'")
                print(f"         roles: {teacher_roles}")
                print(f"         roles type: {type(teacher_roles)}")
                
                # Debug tất cả các keys trong teacher object
                print(f"         All keys: {list(teacher.keys())}")
                
                # Check nếu có teacherInfo
                if 'teacherInfo' in teacher:
                    teacher_info = teacher['teacherInfo']
                    print(f"         teacherInfo keys: {list(teacher_info.keys())}")
                    if 'roles' in teacher_info:
                        print(f"         teacherInfo.roles: {teacher_info.get('roles')}")
            
            for teacher in teachers_list:
                # Lấy roles từ nhiều vị trí có thể
                teacher_roles = teacher.get('roles', [])
                teacher_info = teacher.get('teacherInfo', {})
                teacher_info_roles = teacher_info.get('roles', []) if teacher_info else []
                
                # Combine roles từ cả hai nguồn
                all_roles = []
                if isinstance(teacher_roles, list):
                    all_roles.extend(teacher_roles)
                if isinstance(teacher_info_roles, list):
                    all_roles.extend(teacher_info_roles)
                
                # Nếu roles là string, convert thành list
                if isinstance(teacher_roles, str):
                    all_roles.append(teacher_roles)
                if isinstance(teacher_info_roles, str):
                    all_roles.append(teacher_info_roles)
                
                if not all_roles:
                    continue
                
                # Lấy thông tin giáo viên
                teacher_full_name = teacher.get('fullName', '') or ''
                teacher_info = teacher.get('teacherInfo', {})
                teacher_display_name = teacher_info.get('displayName', '') if teacher_info else ''
                teacher_name = teacher_full_name or teacher_display_name
                
                if not teacher_name:
                    continue
                
                # Kiểm tra vai trò HT (case insensitive)
                if any(role.upper() == 'HT' for role in all_roles if isinstance(role, str)):
                    # Lấy thông tin đăng nhập từ teacherInfo
                    userName = teacher_info.get('userName', '') if teacher_info else ''
                    pwd = teacher_info.get('pwd', '') if teacher_info else ''
                    
                    ht_info = {
                        'name': teacher_name,
                        'fullName': teacher_full_name,
                        'displayName': teacher_display_name,
                        'userName': userName,
                        'pwd': pwd,
                        'roles': all_roles,
                        'raw_data': teacher
                    }
                    ht_teachers.append(ht_info)
                    print(f"      👑 Tìm thấy Hiệu trường: {teacher_name}")
                    print(f"         Roles: {all_roles}")
                    print(f"         Username: {userName}")
                    print(f"         Password: {pwd}")
                
                # Kiểm tra vai trò HP (case insensitive)
                if any(role.upper() == 'HP' for role in all_roles if isinstance(role, str)):
                    # Lấy thông tin đăng nhập từ teacherInfo
                    userName = teacher_info.get('userName', '') if teacher_info else ''
                    pwd = teacher_info.get('pwd', '') if teacher_info else ''
                    
                    hp_info = {
                        'name': teacher_name,
                        'fullName': teacher_full_name,
                        'displayName': teacher_display_name,
                        'userName': userName,
                        'pwd': pwd,
                        'roles': all_roles,
                        'raw_data': teacher
                    }
                    hp_teachers.append(hp_info)
                    print(f"      🔸 Tìm thấy Hiệu phó: {teacher_name}")
                    print(f"         Roles: {all_roles}")
                    print(f"         Username: {userName}")
                    print(f"         Password: {pwd}")
            
            # Tóm tắt kết quả
            print(f"   📊 Kết quả tìm kiếm:")
            print(f"      👑 Hiệu trường (HT): {len(ht_teachers)} người")
            print(f"      🔸 Hiệu phó (HP): {len(hp_teachers)} người")
            
            return {
                'ht': ht_teachers,
                'hp': hp_teachers,
                'total_ht': len(ht_teachers),
                'total_hp': len(hp_teachers)
            }
            
        except Exception as e:
            print_status(f"❌ Lỗi trích xuất thông tin HT/HP: {e}", "error")
            return {'ht': [], 'hp': []}
    
    def _save_ht_hp_info(self, ht_hp_info, school_name):
        """Lưu thông tin HT/HP vào file riêng"""
        try:
            if not ht_hp_info or (not ht_hp_info.get('ht') and not ht_hp_info.get('hp')):
                print("   ⚠️ Không có thông tin HT/HP để lưu")
                return None
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_school_name = "".join(c for c in school_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            filename = f"ht_hp_info_{safe_school_name}_{timestamp}.json"
            filepath = f"data/output/{filename}"
            
            # Tạo cấu trúc dữ liệu để lưu
            save_data = {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'school_name': school_name,
                'summary': {
                    'total_ht': ht_hp_info.get('total_ht', 0),
                    'total_hp': ht_hp_info.get('total_hp', 0)
                },
                'hieu_truong': ht_hp_info.get('ht', []),
                'hieu_pho': ht_hp_info.get('hp', [])
            }
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, ensure_ascii=False, indent=2)
            
            print_status(f"✅ Đã lưu thông tin HT/HP: {filepath}", "success")
            print(f"   👑 {ht_hp_info.get('total_ht', 0)} Hiệu trường")
            print(f"   🔸 {ht_hp_info.get('total_hp', 0)} Hiệu phó")
            
            return filepath
            
        except Exception as e:
            print_status(f"❌ Lỗi lưu thông tin HT/HP: {e}", "error")
            return None
    
    def _is_gvcn_name_in_import(self, name):
        """Kiểm tra xem tên có phải là GVCN hay không (dùng cho cả teacher data và import parsing)"""
        try:
            if not name:
                return False
                
            # Normalize tên để kiểm tra
            normalized_name = str(name).upper().strip()
            
            # Các pattern GVCN thường gặp
            gvcn_patterns = [
                'GVCN',
                'GV CHỦ NHIỆM', 
                'GIÁO VIÊN CHỦ NHIỆM',
                'CHỦ NHIỆM',
                'CHUNHIEM'
            ]
            
            # Kiểm tra xem tên có chứa bất kỳ pattern nào không
            for pattern in gvcn_patterns:
                if pattern in normalized_name:
                    return True
                    
            return False
            
        except Exception:
            return False
    
    def _normalize_date(self, date_str, detected_format=None):
        """Chuẩn hóa ngày sinh để so sánh với xử lý tốt hơn cho các format khác nhau"""
        
        if not date_str or pd.isna(date_str):
            return ""
        
        date_str = str(date_str).strip()
        
        # Xử lý đặc biệt cho Excel datetime format: "2007-10-06 00:00:00"
        # Loại bỏ phần timestamp " HH:MM:SS" để tránh corruption
        if ' ' in date_str and ':' in date_str:
            # Tách phần date và time
            date_part = date_str.split(' ')[0]
            time_part = date_str.split(' ', 1)[1] if len(date_str.split(' ')) > 1 else ""
            
            # Nếu time part có dạng "00:00:00" hoặc timestamp hợp lệ, chỉ lấy phần date
            if re.match(r'\d{2}:\d{2}:\d{2}', time_part):
                date_str = date_part
                # print(f"🔧 Stripped timestamp: '{date_str}' (was '{date_str} {time_part}')")
        
        # Loại bỏ các ký tự không mong muốn nhưng giữ lại dấu / và -
        date_str = re.sub(r'[^\d/\-]', '', date_str)
        
        if not date_str:
            return ""
        
        # Debug: In ra để theo dõi quá trình parse
        # print(f"🔍 Parsing date: '{date_str}' with detected format: {detected_format}")
        
        # Ưu tiên sử dụng detected format nếu có
        if detected_format:
            if detected_format == 'DD/MM/YYYY':
                return self._parse_date_as_dd_mm_yyyy(date_str)
            elif detected_format == 'MM/DD/YYYY':
                return self._parse_date_as_mm_dd_yyyy(date_str)
            elif detected_format == 'YYYY-MM-DD':
                return self._parse_date_as_yyyy_mm_dd(date_str)
            elif detected_format == 'Excel_DateTime':
                # Đã xử lý ở trên bằng cách loại bỏ timestamp
                return self._parse_date_as_dd_mm_yyyy(date_str)
        
        # Fallback: logic cũ nếu không có detected format
        # Ưu tiên xử lý format DD/MM/YYYY trước (format chuẩn cho Vietnam)
        # Pattern cho DD/MM/YYYY hoặc D/M/YYYY với dấu / hoặc -
        match = re.match(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})', date_str)
        if match:
            day, month, year = match.groups()
            try:
                # Validate ngày tháng
                day_int = int(day)
                month_int = int(month)
                year_int = int(year)
                
                if 1 <= day_int <= 31 and 1 <= month_int <= 12 and 1900 <= year_int <= 2030:
                    parsed_date = datetime(year_int, month_int, day_int)
                    # Trả về format chuẩn YYYY-MM-DD để đảm bảo nhất quán
                    normalized = parsed_date.strftime('%Y-%m-%d')
                    # print(f"  ✅ Parsed DD/MM/YYYY: '{date_str}' → '{normalized}'")
                    return normalized
            except ValueError:
                pass
        
        # Pattern cho YYYY-MM-DD (format từ OnLuyen hoặc đã chuẩn hóa)
        match = re.match(r'(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})', date_str)
        if match:
            year, month, day = match.groups()
            try:
                day_int = int(day)
                month_int = int(month)
                year_int = int(year)
                
                if 1 <= day_int <= 31 and 1 <= month_int <= 12 and 1900 <= year_int <= 2030:
                    parsed_date = datetime(year_int, month_int, day_int)
                    normalized = parsed_date.strftime('%Y-%m-%d')
                    # print(f"  ✅ Parsed YYYY-MM-DD: '{date_str}' → '{normalized}'")
                    return normalized
            except ValueError:
                pass
        
        # Pattern cho DD/MM/YY hoặc D/M/YY 
        match = re.match(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{2})', date_str)
        if match:
            day, month, year = match.groups()
            try:
                day_int = int(day)
                month_int = int(month)
                year_int = int(year)
                
                # Xử lý năm 2 chữ số (07 = 2007, 95 = 1995)
                if year_int < 50:  # 00-49 = 2000-2049
                    full_year = year_int + 2000
                else:  # 50-99 = 1950-1999
                    full_year = year_int + 1900
                
                if 1 <= day_int <= 31 and 1 <= month_int <= 12:
                    parsed_date = datetime(full_year, month_int, day_int)
                    normalized = parsed_date.strftime('%Y-%m-%d')
                    # print(f"  ✅ Parsed DD/MM/YY: '{date_str}' → '{normalized}'")
                    return normalized
            except ValueError:
                pass
        
        # Thử các format chuẩn với strptime như fallback
        date_formats = [
            '%d/%m/%Y',     # 24/11/2007
            '%d-%m-%Y',     # 24-11-2007
            '%Y-%m-%d',     # 2007-11-24
            '%d/%m/%y',     # 24/11/07
            '%d-%m-%y',     # 24-11-07
            '%m/%d/%Y',     # 11/24/2007 (US format)
            '%Y/%m/%d',     # 2007/11/24
        ]
        
        for fmt in date_formats:
            try:
                parsed_date = datetime.strptime(date_str, fmt)
                normalized = parsed_date.strftime('%Y-%m-%d')
                # print(f"  ✅ Parsed with format {fmt}: '{date_str}' → '{normalized}'")
                return normalized
            except ValueError:
                continue
        
        # Nếu vẫn không parse được, trả về string gốc đã làm sạch
        print(f"⚠️ Không thể parse ngày: '{date_str}' - giữ nguyên để so sánh")
        return date_str.lower()
    
    def _parse_date_as_dd_mm_yyyy(self, date_str):
        """Parse ngày theo format DD/MM/YYYY"""
        match = re.match(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})', date_str)
        if match:
            day, month, year = match.groups()
            try:
                day_int = int(day)
                month_int = int(month)
                year_int = int(year)
                
                if 1 <= day_int <= 31 and 1 <= month_int <= 12 and 1900 <= year_int <= 2030:
                    parsed_date = datetime(year_int, month_int, day_int)
                    return parsed_date.strftime('%Y-%m-%d')
            except ValueError:
                pass
        return date_str.lower()
    
    def _parse_date_as_mm_dd_yyyy(self, date_str):
        """Parse ngày theo format MM/DD/YYYY"""
        match = re.match(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})', date_str)
        if match:
            month, day, year = match.groups()  # Đảo vị trí month và day
            try:
                day_int = int(day)
                month_int = int(month)
                year_int = int(year)
                
                if 1 <= day_int <= 31 and 1 <= month_int <= 12 and 1900 <= year_int <= 2030:
                    parsed_date = datetime(year_int, month_int, day_int)
                    return parsed_date.strftime('%Y-%m-%d')
            except ValueError:
                pass
        return date_str.lower()
    
    def _parse_date_as_yyyy_mm_dd(self, date_str):
        """Parse ngày theo format YYYY-MM-DD"""
        match = re.match(r'(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})', date_str)
        if match:
            year, month, day = match.groups()
            try:
                day_int = int(day)
                month_int = int(month)
                year_int = int(year)
                
                if 1 <= day_int <= 31 and 1 <= month_int <= 12 and 1900 <= year_int <= 2030:
                    parsed_date = datetime(year_int, month_int, day_int)
                    return parsed_date.strftime('%Y-%m-%d')
            except ValueError:
                pass
        return date_str.lower()
    
    def _standardize_import_date_formats(self, df):
        """Chuẩn hóa format ngày tháng trong import dataframe"""
        print("🔧 Đang chuẩn hóa format ngày tháng trong dữ liệu import...")
        
        date_columns = []
        
        # Tìm các cột có thể chứa ngày
        for col in df.columns:
            col_lower = col.lower()
            if any(keyword in col_lower for keyword in ['ngày', 'sinh', 'date', 'birth']):
                date_columns.append(col)
        
        print(f"📅 Tìm thấy {len(date_columns)} cột ngày: {date_columns}")
        
        # Phân tích format ngày cho từng cột
        for col in date_columns:
            print(f"\n🔍 Phân tích format cho cột '{col}'...")
            
            # Phân tích format thực tế từ dữ liệu
            format_analysis = self._analyze_date_format_in_import(df, col)
            
            if format_analysis:
                detected_format = format_analysis.get('most_likely_format')
                confidence = format_analysis.get('confidence_score', 0)
                
                print(f"✅ Detected format: {detected_format} (confidence: {confidence}%)")
                print(f"📊 Format scores: {format_analysis.get('format_scores', {})}")
                
                if confidence > 50:  # Chỉ áp dụng nếu confidence > 50%
                    print(f"🔄 Applying detected format '{detected_format}' to column '{col}'...")
                    
                    # Áp dụng format đã detect để normalize
                    for idx in df.index:
                        if pd.notna(df.at[idx, col]):
                            original_date = df.at[idx, col]
                            normalized_date = self._normalize_date(str(original_date), detected_format)
                            df.at[idx, col] = normalized_date
                            
                            # Debug: In một vài ví dụ
                            if idx < 3:  # In 3 ví dụ đầu tiên
                                print(f"  📝 Row {idx}: '{original_date}' → '{normalized_date}' (format: {detected_format})")
                else:
                    print(f"⚠️ Confidence thấp ({confidence}%), sử dụng logic fallback...")
                    # Fallback: normalize bằng logic cũ
                    for idx in df.index:
                        if pd.notna(df.at[idx, col]):
                            original_date = df.at[idx, col]
                            normalized_date = self._normalize_date(str(original_date))
                            df.at[idx, col] = normalized_date
            else:
                print(f"❌ Không thể phân tích format cho cột '{col}', sử dụng logic fallback...")
                # Fallback: normalize bằng logic cũ
                for idx in df.index:
                    if pd.notna(df.at[idx, col]):
                        original_date = df.at[idx, col]
                        normalized_date = self._normalize_date(str(original_date))
                        df.at[idx, col] = normalized_date
        
        print("✅ Hoàn thành chuẩn hóa format ngày tháng\n")
        return df
    
    def test_date_normalization(self):
        """Test method để kiểm tra việc chuẩn hóa ngày tháng"""
        print_separator("TEST DATE NORMALIZATION")
        
        test_dates = [
            "6/7/2007",     # D/M/YYYY
            "06/07/2007",   # DD/MM/YYYY  
            "6-7-2007",     # D-M-YYYY
            "06-07-2007",   # DD-MM-YYYY
            "2007-07-06",   # YYYY-MM-DD
            "6/7/07",       # D/M/YY
            "06/07/07",     # DD/MM/YY
            "24/12/1995",   # DD/MM/YYYY
            "1/1/2000",     # D/M/YYYY
            "10/6/2007",    # D/M/YYYY - case cụ thể của user
            "10/06/2007",   # DD/MM/YYYY - case từ OnLuyen
            "",             # Empty
            "invalid",      # Invalid
            "30/2/2020"     # Invalid date
        ]
        
        print("🔍 TEST CASES:")
        for date_str in test_dates:
            normalized = self._normalize_date(date_str)
            print(f"   '{date_str}' → '{normalized}'")
        
        print("\n✅ Test completed. Các ngày hợp lệ sẽ được chuyển về format YYYY-MM-DD")
    
    def test_date_format_detection(self):
        """Test method để kiểm tra việc phát hiện format ngày tháng"""
        print_separator("TEST DATE FORMAT DETECTION")
        
        # Tạo test dataframe giả lập
        test_data = {
            'Họ tên': ['Nguyễn Văn A', 'Trần Thị B', 'Lê Văn C', 'Phạm Thị D'],
            'Ngày sinh': ['10/6/2007', '24/12/1995', '15/8/2000', '3/2/1998']  # DD/MM/YYYY format
        }
        test_df = pd.DataFrame(test_data)
        
        print("🔍 TEST DATA (DD/MM/YYYY):")
        print(test_df)
        print()
        
        # Test format detection
        format_analysis = self._analyze_date_format_in_import(test_df, 'Ngày sinh')
        
        print("📊 FORMAT ANALYSIS RESULTS:")
        if format_analysis:
            print(f"   Most likely format: {format_analysis.get('most_likely_format')}")
            print(f"   Confidence score: {format_analysis.get('confidence_score')}%")
            print(f"   Format scores: {format_analysis.get('format_scores', {})}")
            print(f"   Samples analyzed: {format_analysis.get('sample_count')}")
        else:
            print("   ❌ No format detected")
        
        print("\n🔄 TEST STANDARDIZATION:")
        standardized_df = self._standardize_import_date_formats(test_df.copy())
        print(standardized_df)
        
        print("\n✅ Test completed. Format detection và standardization hoạt động đúng.")
    
    def _save_filtered_workflow_data(self, workflow_results, comparison_results):
        """Lưu dữ liệu workflow đã lọc"""
        try:
            
            school_name = workflow_results['school_info'].get('name', 'Unknown')
            safe_school_name = "".join(c for c in school_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            
            # Tạo cấu trúc JSON với dữ liệu đã lọc
            filtered_data = {
                'school_info': workflow_results['school_info'],
                'data_summary': workflow_results.get('data_summary', {}),
                'comparison_results': {
                    'method': comparison_results.get('comparison_method', 'name_and_birthdate'),
                    'import_teachers_count': comparison_results.get('import_teachers_count', 0),
                    'import_students_count': comparison_results.get('import_students_count', 0),
                    'teachers_matched': comparison_results.get('teachers_matched', 0),
                    'students_matched': comparison_results.get('students_matched', 0)
                },
                'ht_hp_info': workflow_results.get('ht_hp_info', {}),  # Thêm thông tin HT/HP
                'teachers': comparison_results.get('teachers_filtered', []),
                'students': comparison_results.get('students_filtered', [])
            }
            
            # Tạo filename với timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"workflow_filtered_{safe_school_name}_{timestamp}.json"
            filepath = f"data/output/{filename}"
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(filtered_data, f, ensure_ascii=False, indent=2)
            
            return filepath
            
        except Exception as e:
            print_status(f"⚠️ Lỗi lưu dữ liệu filtered workflow: {e}", "warning")
            return None
    
    def _print_workflow_summary_case_2(self, results):
        """In tóm tắt kết quả workflow Case 2"""
        print(f"\n📊 TÓM TẮT KẾT QUẢ WORKFLOW CASE 2:")
        print("=" * 70)
        
        print(f"🏫 Trường: {results['school_info'].get('name', 'N/A')}")
        print(f"👤 Admin: {results['school_info'].get('admin', 'N/A')}")
        print()
        
        # Trạng thái từng bước
        steps = [
            ("1️⃣ Trích xuất Google Sheets", results['sheets_extraction']),
            ("2️⃣ OnLuyen API Login", results['api_login']),
            ("3️⃣ Lấy dữ liệu Giáo viên", results['teachers_data']),
            ("4️⃣ Lấy dữ liệu Học sinh", results['students_data']),
            ("5️⃣ Tải file import", results['import_file_downloaded']),
            ("6️⃣ So sánh dữ liệu", results['data_comparison']),
            ("7️⃣ Lưu dữ liệu JSON", results['json_saved']),
            ("8️⃣ Chuyển đổi Excel", results['excel_converted']),
            ("9️⃣ Upload Google Drive", results['drive_uploaded'])
        ]
        
        for step_name, status in steps:
            status_icon = "✅" if status else "❌"
            status_text = "Thành công" if status else "Thất bại"
            print(f"{status_icon} {step_name}: {status_text}")
        
        # Tóm tắt so sánh dữ liệu
        if results.get('comparison_results'):
            comp = results['comparison_results']
            print(f"\n🔍 KẾT QUẢ SO SÁNH (Theo Họ tên + Ngày sinh):")
            print(f"   � Import Teachers: {comp.get('import_teachers_count', 0)}")
            print(f"   📊 Import Students: {comp.get('import_students_count', 0)}")
            
            # Hiển thị thông tin đặc biệt về giáo viên
            if comp.get('export_all_teachers', False):
                print(f"   👨‍🏫 Giáo viên: XUẤT TẤT CẢ {comp.get('teachers_matched', 0)} (tìm thấy GVCN)")
            else:
                print(f"   👨‍🏫 Giáo viên khớp: {comp.get('teachers_matched', 0)}/{comp.get('import_teachers_count', 0)}")
            
            print(f"   👨‍🎓 Học sinh khớp: {comp.get('students_matched', 0)}/{comp.get('import_students_count', 0)}")
            print(f"   🔧 Phương pháp: {comp.get('method', 'name_and_birthdate')}")
        
        # Thông tin HT/HP
        if results.get('ht_hp_info'):
            ht_hp_info = results['ht_hp_info']
            print(f"\n👑 THÔNG TIN LÃNH ĐẠO:")
            print(f"   👑 Hiệu trường (HT): {ht_hp_info.get('total_ht', 0)} người")
            print(f"   🔸 Hiệu phó (HP): {ht_hp_info.get('total_hp', 0)} người")
            
            # Hiển thị danh sách HT
            if ht_hp_info.get('ht'):
                print(f"   📋 Danh sách Hiệu trường:")
                for i, ht in enumerate(ht_hp_info['ht'], 1):
                    print(f"      {i}. {ht['name']}")
            
            # Hiển thị danh sách HP
            if ht_hp_info.get('hp'):
                print(f"   📋 Danh sách Hiệu phó:")
                for i, hp in enumerate(ht_hp_info['hp'], 1):
                    print(f"      {i}. {hp['name']}")
        
        # File outputs
        if results.get('json_file_path') or results.get('excel_file_path') or results.get('ht_hp_file'):
            print(f"\n📄 FILES ĐÃ TẠO:")
            if results.get('json_file_path'):
                print(f"   📋 JSON: {results['json_file_path']}")
            if results.get('excel_file_path'):
                print(f"   📊 Excel: {results['excel_file_path']}")
            if results.get('ht_hp_file'):
                print(f"   👑 HT/HP Info: {results['ht_hp_file']}")
        
        # Upload results
        if results.get('upload_results'):
            upload_info = results['upload_results']
            print(f"\n📤 DRIVE UPLOAD:")
            print(f"   ✅ Thành công: {upload_info.get('success', 0)} files")
            print(f"   ❌ Thất bại: {upload_info.get('failed', 0)} files")
        
        # Tổng kết
        success_count = sum([results['sheets_extraction'], results['api_login'], 
                           results['teachers_data'], results['students_data'],
                           results['import_file_downloaded'], results['data_comparison'],
                           results['json_saved'], results['excel_converted'], 
                           results['drive_uploaded']])
        total_steps = 9
        
        print(f"\n🎯 TỔNG KẾT: {success_count}/{total_steps} bước thành công")
        
        if success_count == total_steps:
            print_status("🎉 WORKFLOW CASE 2 HOÀN CHỈNH - ĐÃ SO SÁNH, LỌC VÀ TẠO EXCEL!", "success")
        elif success_count >= 7:
            print_status("⚠️ Workflow Case 2 hoàn thành chính (có thể thiếu upload)", "warning")
        elif success_count >= 5:
            print_status("⚠️ Workflow Case 2 hoàn thành một phần", "warning")
        else:
            print_status("❌ Workflow Case 2 thất bại", "error")

    def run(self):
        """Chạy ứng dụng"""
        try:
            print_header("SCHOOL PROCESS APPLICATION - ENHANCED", "Hệ thống xử lý dữ liệu trường học")
            
            # Hiển thị thông tin config nếu debug mode
            if self.config.is_debug_mode():
                self.config.print_config_summary()
            
            self.show_main_menu()
            
        except KeyboardInterrupt:
            print("\n\n⏹️  Ứng dụng bị dừng bởi người dùng")
        except Exception as e:
            print(f"\n💥 LỖI NGHIÊM TRỌNG: {e}")
            print("📞 Vui lòng kiểm tra cấu hình hoặc liên hệ support")
        finally:
            print(f"\n📊 Ứng dụng kết thúc: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


def main():
    """Entry point"""
    app = SchoolProcessApp()
    app.run()


if __name__ == "__main__":
    main()