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
from pathlib import Path

# Thêm project root vào Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from config.config_manager import get_config
from utils.menu_utils import *
from utils.file_utils import ensure_directories


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
            from processors.local_processor import LocalDataProcessor
            
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
            "Tích hợp hoàn chỉnh: Sheets → Login → Dữ liệu → Excel",
            "Lấy danh sách Giáo viên",
            "Lấy danh sách Học sinh"
        ]
        
        handlers = [
            self.onluyen_complete_workflow,
            self.onluyen_get_teachers,
            self.onluyen_get_students
        ]
        
        run_menu_loop("ONLUYEN API INTEGRATION", options, handlers)
    
    def onluyen_get_teachers(self):
        """Lấy danh sách giáo viên"""
        print_separator("LẤY DANH SÁCH GIÁO VIÊN")
        
        try:
            from config.onluyen_api import OnLuyenAPIClient
            
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
            from config.onluyen_api import OnLuyenAPIClient
            
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
            
            from extractors import GoogleSheetsExtractor
            
            extractor = GoogleSheetsExtractor()
            sheet_name = get_user_input("Nhập tên sheet (mặc định: ED-2025)") or "ED-2025"
            
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
            
            from config.onluyen_api import OnLuyenAPIClient
            
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

            from datetime import datetime
            
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
            print_status("BƯỚC 1: Trích xuất dữ liệu từ Google Sheets", "info")
            
            from extractors import GoogleSheetsExtractor
            
            extractor = GoogleSheetsExtractor()
            sheet_name = get_user_input("Nhập tên sheet (mặc định: ED-2025)") or "ED-2025"
            
            print_status(f"Đang trích xuất dữ liệu từ sheet: {sheet_name}", "info")
            school_data = extractor.extract_school_data(sheet_name=sheet_name)
            
            if not school_data:
                print_status("❌ Không thể trích xuất dữ liệu từ Google Sheets", "error")
                return
            
            workflow_results['sheets_extraction'] = True
            print_status(f"✅ Đã trích xuất {len(school_data)} trường học", "success")
            
            # Chọn trường để xử lý
            if len(school_data) == 1:
                selected_school = school_data[0]
                print_status("Tự động chọn trường duy nhất", "info")
            else:
                print("\n📋 DANH SÁCH TRƯỜNG ĐÃ TRÍCH XUẤT:")
                for i, school in enumerate(school_data, 1):
                    school_name = school.get('Tên trường', 'N/A')
                    admin_email = school.get('Admin', 'N/A')
                    print(f"   {i}. {school_name} (Admin: {admin_email})")
                
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
            
            # Lấy thông tin trường
            school_name = selected_school.get('Tên trường', 'N/A')
            admin_email = selected_school.get('Admin', '')
            password = selected_school.get('Mật khẩu', '')
            drive_link = selected_school.get('Link driver dữ liệu', 'N/A')
            
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
            
            from config.onluyen_api import OnLuyenAPIClient
            
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
            
            # Hỏi có muốn mở file Excel không nếu tạo thành công
            if workflow_results['excel_converted'] and workflow_results['excel_file_path']:
                action_options = ["Mở file Excel local"]
                
                if workflow_results['drive_uploaded'] and workflow_results['upload_results'].get('urls'):
                    action_options.append("Mở Google Drive folder")
                
                if len(action_options) > 1:
                    print(f"\n🎯 BẠN CÓ THỂ:")
                    for i, option in enumerate(action_options, 1):
                        print(f"   {i}. {option}")
                    
                    choice = get_user_input(f"Chọn hành động (1-{len(action_options)}, Enter = bỏ qua)")
                    
                    if choice == "1":
                        try:
                            os.startfile(workflow_results['excel_file_path'])
                            print_status("Đã mở file Excel", "success")
                        except Exception as e:
                            print_status(f"Không thể mở file Excel: {e}", "warning")
                    elif choice == "2" and len(action_options) > 1:
                        drive_folder_url = drive_link
                        print_status(f"🔗 Google Drive: {drive_folder_url}", "info")
                        print("💡 Bạn có thể mở link trên trong trình duyệt")
                else:
                    if get_user_confirmation("Bạn có muốn mở file Excel đã tạo?"):
                        try:
                            os.startfile(workflow_results['excel_file_path'])
                            print_status("Đã mở file Excel", "success")
                        except Exception as e:
                            print_status(f"Không thể mở file Excel: {e}", "warning")
            
            # Lưu dữ liệu vào file nếu chưa lưu (fallback)
            if not workflow_results['json_saved'] and (workflow_results['teachers_data'] or workflow_results['students_data']):
                self._save_workflow_data(workflow_results, teachers_result, students_result, password)
            
        except ImportError as e:
            print_status(f"Module không tồn tại: {e}", "error")
        except Exception as e:
            print_status(f"Lỗi trong quy trình tích hợp: {e}", "error")
    
    def _convert_json_to_excel(self, json_file_path):
        """Chuyển đổi file JSON workflow sang Excel"""
        try:
            from converters import JSONToExcelTemplateConverter
            from pathlib import Path
            
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
            import glob
            
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
            from config.google_oauth_drive import GoogleOAuthDriveClient
            
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
            import re
            
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
        
        # File outputs
        if results.get('json_file_path') or results.get('excel_file_path'):
            print(f"\n📄 FILES ĐÃ TẠO:")
            if results.get('json_file_path'):
                print(f"   📋 JSON: {results['json_file_path']}")
            if results.get('excel_file_path'):
                print(f"   📊 Excel: {results['excel_file_path']}")
        
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
            from datetime import datetime
            
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
            from datetime import datetime
            
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
            import json
            from datetime import datetime
            
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
            import glob
            import json
            from pathlib import Path
            
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
            from config.onluyen_api import OnLuyenAPIClient
            
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
            from pathlib import Path
            
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
            from converters import JSONToExcelTemplateConverter
            
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
    
    def _save_teachers_data(self, teachers_list, total_count):
        """Lưu dữ liệu giáo viên vào file JSON"""
        try:
            from datetime import datetime
            
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
            from datetime import datetime
            
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
    
    def _extract_drive_folder_id(self, drive_link):
        """Extract folder ID từ Google Drive link"""
        try:
            import re
            
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
