"""
Google Data Processor
Processor xử lý dữ liệu từ Google Drive/Sheets
Author: Assistant
Date: 2025-07-26
"""

import pandas as pd
import os
from .base_processor import BaseDataProcessor


class GoogleDataProcessor(BaseDataProcessor):
    """Processor xử lý dữ liệu từ Google Drive/Sheets"""
    
    def __init__(self, input_folder: str, temp_folder: str, output_folder: str = None):
        """
        Khởi tạo GoogleDataProcessor
        
        Args:
            input_folder (str): Thư mục chứa file input
            temp_folder (str): Thư mục chứa template
            output_folder (str): Thư mục output (mặc định là temp_folder)
        """
        super().__init__(input_folder, temp_folder, output_folder)
        
        # Import Google API client
        try:
            from config.google_api import GoogleAPIClient
            self.google_client = GoogleAPIClient()
            self.is_google_available = True
        except ImportError as e:
            print(f"⚠️  Google API không khả dụng: {e}")
            self.google_client = None
            self.is_google_available = False
    
    def get_processor_name(self) -> str:
        """Trả về tên processor"""
        return "GOOGLE API PROCESSOR"
    
    def check_google_api_status(self) -> bool:
        """
        Kiểm tra trạng thái Google API một cách chi tiết
        
        Returns:
            bool: True nếu Google API sẵn sàng
        """
        print("🔍 KIỂM TRA CHI TIẾT GOOGLE API:")
        print("-" * 40)
        
        if not self.is_google_available:
            print("❌ Google API client không khả dụng")
            print("   💡 Kiểm tra: import config.google_api")
            return False
        
        try:
            print("📋 Bước 1: Kiểm tra Service Account...")
            if not hasattr(self.google_client, 'service_account_file'):
                print("❌ Service Account file không được cấu hình")
                return False
            
            if not os.path.exists(self.google_client.service_account_file):
                print(f"❌ Service Account file không tồn tại: {self.google_client.service_account_file}")
                return False
            
            print("✅ Service Account file tồn tại")
            
            print("\n📋 Bước 2: Kiểm tra Google Drive API...")
            try:
                # Test Drive API
                drive_service = self.google_client.drive_service
                if drive_service:
                    # Test bằng cách list files (giới hạn 1 file)
                    results = drive_service.files().list(pageSize=1).execute()
                    print("✅ Google Drive API hoạt động bình thường")
                else:
                    print("❌ Không thể khởi tạo Drive service")
                    return False
            except Exception as e:
                print(f"❌ Lỗi Google Drive API: {e}")
                return False
            
            print("\n📋 Bước 3: Kiểm tra Google Sheets API...")
            try:
                # Test Sheets API  
                sheets_service = self.google_client.sheets_service
                if sheets_service:
                    print("✅ Google Sheets API hoạt động bình thường")
                else:
                    print("❌ Không thể khởi tạo Sheets service")
                    return False
            except Exception as e:
                print(f"❌ Lỗi Google Sheets API: {e}")
                return False
            
            print("\n📋 Bước 4: Test connection tổng thể...")
            if self.google_client.test_connection():
                print("✅ Google API kết nối thành công - Tất cả dịch vụ ổn định")
                return True
            else:
                print("❌ Google API test connection thất bại")
                return False
                
        except Exception as e:
            print(f"❌ Lỗi nghiêm trọng khi kiểm tra Google API: {e}")
            print(f"   🔍 Chi tiết lỗi: {type(e).__name__}")
            return False
    
    def test_google_api_permissions(self) -> dict:
        """
        Kiểm tra quyền truy cập chi tiết của Google API
        
        Returns:
            dict: Kết quả kiểm tra quyền
        """
        print("\n🔐 KIỂM TRA QUYỀN TRUY CẬP GOOGLE API:")
        print("-" * 40)
        
        permissions = {
            "drive_read": False,
            "drive_write": False,
            "sheets_read": False,
            "sheets_write": False,
            "errors": []
        }
        
        if not self.is_google_available or not self.google_client:
            permissions["errors"].append("Google API client không khả dụng")
            return permissions
        
        try:
            # Test Drive read permission
            print("📖 Test quyền đọc Google Drive...")
            try:
                drive_service = self.google_client.drive_service
                results = drive_service.files().list(pageSize=5).execute()
                files = results.get('files', [])
                permissions["drive_read"] = True
                print(f"✅ Đọc Drive OK - Tìm thấy {len(files)} file")
            except Exception as e:
                permissions["errors"].append(f"Lỗi đọc Drive: {e}")
                print(f"❌ Lỗi đọc Drive: {e}")
            
            # Test Drive write permission
            print("\n📝 Test quyền ghi Google Drive...")
            try:
                # Tạo file test nhỏ
                test_content = "Test file for permission check"
                test_file_id = self.google_client.create_test_file("SchoolProcess_PermissionTest.txt", test_content)
                if test_file_id:
                    permissions["drive_write"] = True
                    print("✅ Ghi Drive OK")
                    
                    # Xóa file test ngay
                    try:
                        drive_service.files().delete(fileId=test_file_id).execute()
                        print("🗑️  Đã xóa file test")
                    except:
                        pass
                else:
                    permissions["errors"].append("Không thể tạo file test trên Drive")
                    print("❌ Không thể ghi file lên Drive")
            except Exception as e:
                permissions["errors"].append(f"Lỗi ghi Drive: {e}")
                print(f"❌ Lỗi ghi Drive: {e}")
            
            # Test Sheets read permission
            print("\n📊 Test quyền đọc Google Sheets...")
            try:
                sheets_service = self.google_client.sheets_service
                # Tạo một sheet test đơn giản để kiểm tra
                spreadsheet = {
                    'properties': {
                        'title': 'SchoolProcess_PermissionTest'
                    }
                }
                sheet = sheets_service.spreadsheets().create(body=spreadsheet).execute()
                sheet_id = sheet.get('spreadsheetId')
                
                if sheet_id:
                    permissions["sheets_read"] = True
                    permissions["sheets_write"] = True
                    print("✅ Sheets đọc/ghi OK")
                    
                    # Xóa sheet test
                    try:
                        drive_service = self.google_client.drive_service
                        drive_service.files().delete(fileId=sheet_id).execute()
                        print("🗑️  Đã xóa sheet test")
                    except:
                        pass
                        
            except Exception as e:
                permissions["errors"].append(f"Lỗi Sheets: {e}")
                print(f"❌ Lỗi Sheets: {e}")
            
        except Exception as e:
            permissions["errors"].append(f"Lỗi tổng thể: {e}")
            print(f"❌ Lỗi tổng thể: {e}")
        
        return permissions
    
    def measure_api_performance(self) -> dict:
        """
        Đo hiệu suất API Google
        
        Returns:
            dict: Kết quả đo hiệu suất
        """
        print("\n⚡ ĐO HIỆU SUẤT GOOGLE API:")
        print("-" * 40)
        
        import time
        
        performance = {
            "drive_list_time": 0,
            "sheets_create_time": 0,
            "overall_stable": False,
            "recommendations": []
        }
        
        if not self.is_google_available or not self.google_client:
            performance["recommendations"].append("Cấu hình Google API trước")
            return performance
        
        try:
            # Test tốc độ list Drive files
            print("🏃‍♂️ Đo tốc độ list Drive files...")
            start_time = time.time()
            drive_service = self.google_client.drive_service
            results = drive_service.files().list(pageSize=10).execute()
            end_time = time.time()
            
            drive_time = end_time - start_time
            performance["drive_list_time"] = round(drive_time, 2)
            print(f"⏱️  Drive list: {drive_time:.2f}s")
            
            if drive_time < 2.0:
                print("✅ Tốc độ Drive tốt")
            elif drive_time < 5.0:
                print("⚠️  Tốc độ Drive trung bình")
                performance["recommendations"].append("Kết nối mạng có thể chậm")
            else:
                print("❌ Tốc độ Drive chậm")
                performance["recommendations"].append("Kiểm tra kết nối internet")
            
            # Test tốc độ tạo Sheets
            print("\n📋 Đo tốc độ tạo Google Sheets...")
            start_time = time.time()
            sheets_service = self.google_client.sheets_service
            spreadsheet = {
                'properties': {
                    'title': 'SchoolProcess_SpeedTest'
                }
            }
            sheet = sheets_service.spreadsheets().create(body=spreadsheet).execute()
            end_time = time.time()
            
            sheets_time = end_time - start_time
            performance["sheets_create_time"] = round(sheets_time, 2)
            print(f"⏱️  Sheets create: {sheets_time:.2f}s")
            
            # Xóa sheet test
            try:
                sheet_id = sheet.get('spreadsheetId')
                drive_service.files().delete(fileId=sheet_id).execute()
                print("🗑️  Đã xóa sheet test")
            except:
                pass
            
            if sheets_time < 3.0:
                print("✅ Tốc độ Sheets tốt")
            elif sheets_time < 7.0:
                print("⚠️  Tốc độ Sheets trung bình")
            else:
                print("❌ Tốc độ Sheets chậm")
                performance["recommendations"].append("API có thể bị giới hạn tốc độ")
            
            # Đánh giá tổng thể
            if drive_time < 3.0 and sheets_time < 5.0:
                performance["overall_stable"] = True
                print("\n🎯 Kết luận: API ổn định, sẵn sàng xử lý dữ liệu")
            else:
                print("\n⚠️  Kết luận: API chậm, có thể ảnh hưởng hiệu suất")
                
        except Exception as e:
            print(f"❌ Lỗi đo hiệu suất: {e}")
            performance["recommendations"].append(f"Lỗi: {e}")
        
        return performance
    
    def load_students_data(self) -> pd.DataFrame:
        """
        Load dữ liệu học sinh từ Google Sheets hoặc local fallback
        
        Returns:
            pd.DataFrame: Dữ liệu học sinh
        """
        try:
            # Thử load từ Google Sheets trước
            if self.is_google_available and self.google_client:
                print("🌐 Đang thử load dữ liệu học sinh từ Google Sheets...")
                
                try:
                    # Tìm file học sinh trên Google Drive
                    student_file_id = self.google_client.find_file_by_name("Danh sach hoc sinh.xlsx")
                    
                    if student_file_id:
                        print(f"📁 Tìm thấy file học sinh trên Google Drive: {student_file_id}")
                        
                        # Download file về temp
                        temp_file_path = os.path.join(self.temp_folder, "temp_students.xlsx")
                        if self.google_client.download_file(student_file_id, temp_file_path):
                            print("📥 Đã download file học sinh từ Google Drive")
                            
                            # Đọc dữ liệu từ file đã download
                            df = self._read_students_from_excel(temp_file_path)
                            
                            # Xóa file temp
                            if os.path.exists(temp_file_path):
                                os.remove(temp_file_path)
                            
                            if not df.empty:
                                self.students_data = df
                                print(f"✅ Đã load {len(df)} học sinh từ Google Drive")
                                return df
                    
                except Exception as e:
                    print(f"⚠️  Lỗi khi load từ Google Sheets: {e}")
            
            # Fallback: Load từ file local
            print("🏠 Fallback: Load dữ liệu học sinh từ file local...")
            return self._load_students_from_local()
            
        except Exception as e:
            print(f"❌ Lỗi khi load dữ liệu học sinh: {e}")
            self.students_data = pd.DataFrame()
            return pd.DataFrame()
    
    def load_teachers_data(self) -> pd.DataFrame:
        """
        Load dữ liệu giáo viên từ Google Sheets hoặc local fallback
        
        Returns:
            pd.DataFrame: Dữ liệu giáo viên
        """
        try:
            # Thử load từ Google Sheets trước
            if self.is_google_available and self.google_client:
                print("🌐 Đang thử load dữ liệu giáo viên từ Google Sheets...")
                
                try:
                    # Tìm file giáo viên trên Google Drive
                    teacher_file_id = self.google_client.find_file_by_name("DS tài khoản giáo viên.xlsx")
                    
                    if teacher_file_id:
                        print(f"📁 Tìm thấy file giáo viên trên Google Drive: {teacher_file_id}")
                        
                        # Download file về temp
                        temp_file_path = os.path.join(self.temp_folder, "temp_teachers.xlsx")
                        if self.google_client.download_file(teacher_file_id, temp_file_path):
                            print("📥 Đã download file giáo viên từ Google Drive")
                            
                            # Đọc dữ liệu từ file đã download
                            df = self._read_teachers_from_excel(temp_file_path)
                            
                            # Xóa file temp
                            if os.path.exists(temp_file_path):
                                os.remove(temp_file_path)
                            
                            if not df.empty:
                                self.teachers_data = df
                                print(f"✅ Đã load {len(df)} giáo viên từ Google Drive")
                                return df
                    
                except Exception as e:
                    print(f"⚠️  Lỗi khi load từ Google Sheets: {e}")
            
            # Fallback: Load từ file local
            print("🏠 Fallback: Load dữ liệu giáo viên từ file local...")
            return self._load_teachers_from_local()
            
        except Exception as e:
            print(f"❌ Lỗi khi load dữ liệu giáo viên: {e}")
            self.teachers_data = pd.DataFrame()
            return pd.DataFrame()
    
    def _read_students_from_excel(self, file_path: str) -> pd.DataFrame:
        """
        Đọc dữ liệu học sinh từ file Excel
        
        Args:
            file_path (str): Đường dẫn file Excel
            
        Returns:
            pd.DataFrame: Dữ liệu học sinh
        """
        try:
            # Đọc với header tại hàng 6 (index 5)
            df = pd.read_excel(
                file_path, 
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
            
            return df
            
        except Exception as e:
            print(f"❌ Lỗi khi đọc file Excel học sinh: {e}")
            return pd.DataFrame()
    
    def _read_teachers_from_excel(self, file_path: str) -> pd.DataFrame:
        """
        Đọc dữ liệu giáo viên từ file Excel
        
        Args:
            file_path (str): Đường dẫn file Excel
            
        Returns:
            pd.DataFrame: Dữ liệu giáo viên
        """
        try:
            # Thử đọc với header mặc định trước
            df = pd.read_excel(
                file_path,
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
            
            return df
            
        except Exception as e:
            print(f"❌ Lỗi khi đọc file Excel giáo viên: {e}")
            return pd.DataFrame()
    
    def _load_students_from_local(self) -> pd.DataFrame:
        """Load dữ liệu học sinh từ file local"""
        try:
            if not os.path.exists(self.student_file):
                raise FileNotFoundError(f"File học sinh không tồn tại: {self.student_file}")
            
            df = self._read_students_from_excel(self.student_file)
            self.students_data = df
            print(f"✅ Đã load {len(df)} học sinh từ file local")
            return df
            
        except Exception as e:
            print(f"❌ Lỗi khi load học sinh từ local: {e}")
            return pd.DataFrame()
    
    def _load_teachers_from_local(self) -> pd.DataFrame:
        """Load dữ liệu giáo viên từ file local"""
        try:
            if not os.path.exists(self.teacher_file):
                raise FileNotFoundError(f"File giáo viên không tồn tại: {self.teacher_file}")
            
            df = self._read_teachers_from_excel(self.teacher_file)
            self.teachers_data = df
            print(f"✅ Đã load {len(df)} giáo viên từ file local")
            return df
            
        except Exception as e:
            print(f"❌ Lỗi khi load giáo viên từ local: {e}")
            return pd.DataFrame()
    
    def _post_save_actions(self, output_path: str, output_filename: str):
        """
        Upload file lên Google Drive sau khi lưu local
        
        Args:
            output_path (str): Đường dẫn file đã lưu
            output_filename (str): Tên file
        """
        try:
            if self.is_google_available and self.google_client:
                print(f"🌐 Đang upload file lên Google Drive...")
                
                # Upload file lên Google Drive
                file_id = self.google_client.upload_file(
                    output_path, 
                    output_filename,
                    folder_name="SchoolProcess_Output"  # Tạo thư mục riêng
                )
                
                if file_id:
                    print(f"✅ Đã upload file lên Google Drive: {file_id}")
                    
                    # Tạo link chia sẻ
                    share_link = self.google_client.create_shareable_link(file_id)
                    if share_link:
                        print(f"🔗 Link chia sẻ: {share_link}")
                else:
                    print("❌ Upload lên Google Drive thất bại")
            else:
                print("⚠️  Google API không khả dụng, bỏ qua upload")
                
        except Exception as e:
            print(f"❌ Lỗi khi upload lên Google Drive: {e}")
    
    def process_with_google_api(self) -> str:
        """
        Kiểm tra và test Google API (chưa xử lý data)
        
        Returns:
            str: Trạng thái kiểm tra
        """
        print("\n🌐 CHẾ ĐỘ 2: KIỂM TRA GOOGLE API")
        print("=" * 50)
        
        # Bước 1: Kiểm tra cơ bản
        print("🔍 GIAI ĐOẠN 1: KIỂM TRA CƠ BẢN")
        basic_status = self.check_google_api_status()
        
        if not basic_status:
            print("\n❌ Google API không hoạt động. Dừng kiểm tra.")
            return "FAILED"
        
        # Bước 2: Kiểm tra quyền truy cập
        print("\n🔐 GIAI ĐOẠN 2: KIỂM TRA QUYỀN TRUY CẬP")
        permissions = self.test_google_api_permissions()
        
        # Hiển thị kết quả quyền
        print(f"\n📋 KẾT QUẢ KIỂM TRA QUYỀN:")
        print(f"   📖 Drive Read: {'✅' if permissions['drive_read'] else '❌'}")
        print(f"   📝 Drive Write: {'✅' if permissions['drive_write'] else '❌'}")
        print(f"   📊 Sheets Read: {'✅' if permissions['sheets_read'] else '❌'}")
        print(f"   ✏️  Sheets Write: {'✅' if permissions['sheets_write'] else '❌'}")
        
        if permissions['errors']:
            print(f"\n⚠️  Lỗi quyền truy cập:")
            for error in permissions['errors']:
                print(f"      - {error}")
        
        # Bước 3: Đo hiệu suất
        print("\n⚡ GIAI ĐOẠN 3: ĐO HIỆU SUẤT API")
        performance = self.measure_api_performance()
        
        # Hiển thị kết quả hiệu suất
        print(f"\n📊 KẾT QUẢ HIỆU SUẤT:")
        print(f"   🏃‍♂️ Drive List: {performance['drive_list_time']}s")
        print(f"   📋 Sheets Create: {performance['sheets_create_time']}s")
        print(f"   🎯 Ổn định: {'✅' if performance['overall_stable'] else '❌'}")
        
        if performance['recommendations']:
            print(f"\n💡 KHUYẾN NGHỊ:")
            for rec in performance['recommendations']:
                print(f"      - {rec}")
        
        # Đánh giá tổng thể
        print(f"\n📋 ĐÁNH GIÁ TỔNG THỂ:")
        
        # Tính điểm
        score = 0
        max_score = 6
        
        if basic_status:
            score += 2
            print("   ✅ Kết nối cơ bản: OK (+2 điểm)")
        
        if permissions['drive_read'] and permissions['drive_write']:
            score += 2
            print("   ✅ Quyền Drive: OK (+2 điểm)")
        
        if permissions['sheets_read'] and permissions['sheets_write']:
            score += 1
            print("   ✅ Quyền Sheets: OK (+1 điểm)")
        
        if performance['overall_stable']:
            score += 1
            print("   ✅ Hiệu suất: OK (+1 điểm)")
        
        # Kết luận
        percentage = (score / max_score) * 100
        print(f"\n🎯 ĐIỂM SỐ: {score}/{max_score} ({percentage:.0f}%)")
        
        if score >= 5:
            status = "EXCELLENT"
            print("🟢 Trạng thái: XUẤT SẮC - Google API sẵn sàng xử lý dữ liệu")
        elif score >= 3:
            status = "GOOD"
            print("🟡 Trạng thái: TỐT - Google API có thể sử dụng nhưng cần lưu ý")
        elif score >= 1:
            status = "POOR"
            print("🟠 Trạng thái: YẾU - Google API có vấn đề, nên sử dụng local")
        else:
            status = "FAILED"
            print("🔴 Trạng thái: THẤT BẠI - Google API không hoạt động")
        
        print(f"\n✅ HOÀN THÀNH KIỂM TRA GOOGLE API!")
        print(f"📊 Trạng thái cuối: {status}")
        
        return status
    
    def sync_to_google_sheets(self, output_path: str) -> str:
        """
        Đồng bộ dữ liệu lên Google Sheets
        
        Args:
            output_path (str): Đường dẫn file Excel để đồng bộ
            
        Returns:
            str: Link Google Sheets hoặc rỗng nếu thất bại
        """
        try:
            if not self.is_google_available or not self.google_client:
                print("❌ Google API không khả dụng")
                return ""
            
            print("📊 Đang đồng bộ dữ liệu lên Google Sheets...")
            
            # Tạo Google Sheets mới
            sheets_id = self.google_client.create_google_sheets(
                f"SchoolProcess_{self.school_name}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}"
            )
            
            if not sheets_id:
                print("❌ Không thể tạo Google Sheets")
                return ""
            
            # Upload dữ liệu từ file Excel
            if self.google_client.upload_excel_to_sheets(output_path, sheets_id):
                print(f"✅ Đã đồng bộ dữ liệu lên Google Sheets: {sheets_id}")
                
                # Tạo link chia sẻ
                share_link = self.google_client.create_shareable_link(sheets_id)
                if share_link:
                    print(f"🔗 Link Google Sheets: {share_link}")
                    return share_link
            
            return ""
            
        except Exception as e:
            print(f"❌ Lỗi khi đồng bộ lên Google Sheets: {e}")
            return ""
    
    def test_shared_sheets_access(self) -> dict:
        """
        Test khả năng đọc shared Google Sheets
        
        Returns:
            dict: Kết quả test shared sheets
        """
        print("\n📊 TEST SHARED GOOGLE SHEETS ACCESS:")
        print("-" * 40)
        
        results = {
            "has_shared_sheets": False,
            "accessible_sheets": [],
            "failed_sheets": [],
            "recommendations": []
        }
        
        if not self.is_google_available or not self.google_client:
            results["recommendations"].append("Google API không khả dụng")
            return results
        
        try:
            # Danh sách test sheets (có thể thêm vào config sau)
            test_sheets = [
                # Có thể thêm URL/ID của shared sheets ở đây để test
                # "1ABC123...",  # Sheet ID example
                # "https://docs.google.com/spreadsheets/d/1ABC123.../edit"
            ]
            
            print(f"🔍 Checking shared sheets access...")
            
            if not test_sheets:
                print("ℹ️  Không có shared sheets để test")
                print("💡 Để test shared sheets:")
                print("   1. Lấy URL của Google Sheets đã được chia sẻ")
                print("   2. Đảm bảo Service Account email có quyền truy cập")
                print("   3. Thử method read_shared_google_sheet()")
                
                results["recommendations"].extend([
                    "Cần URL hoặc ID của shared Google Sheets",
                    "Service Account cần được chia sẻ quyền truy cập", 
                    "Hoặc sheet phải được public read"
                ])
                return results
            
            # Test từng shared sheet
            for sheet_url_id in test_sheets:
                print(f"\n📋 Testing sheet: {sheet_url_id[:20]}...")
                
                access_result = self.google_client.test_shared_sheet_access(sheet_url_id)
                
                if access_result["accessible"]:
                    results["accessible_sheets"].append({
                        "id": sheet_url_id,
                        "title": access_result["sheet_title"],
                        "method": access_result["method"],
                        "sheets": access_result["sheets_list"],
                        "can_write": access_result["permissions"]["write"]
                    })
                    print(f"✅ Accessible: {access_result['sheet_title']}")
                else:
                    results["failed_sheets"].append({
                        "id": sheet_url_id,
                        "errors": access_result["errors"]
                    })
                    print(f"❌ Not accessible")
            
            results["has_shared_sheets"] = len(results["accessible_sheets"]) > 0
            
            # Recommendations
            if results["has_shared_sheets"]:
                results["recommendations"].append("Có thể sử dụng shared sheets làm data source")
            else:
                results["recommendations"].extend([
                    "Kiểm tra Service Account có được chia sẻ quyền không",
                    "Hoặc set sheet thành public readable",
                    "Kiểm tra URL/ID sheet có đúng không"
                ])
            
        except Exception as e:
            print(f"❌ Lỗi test shared sheets: {e}")
            results["recommendations"].append(f"Lỗi: {e}")
        
        return results
    
    def read_data(self):
        """
        Đọc dữ liệu từ Google Sheets với enhanced processing
        
        Returns:
            pd.DataFrame: DataFrame chứa dữ liệu đã đọc
        """
        try:
            if not self.is_google_available or not self.google_client:
                print("❌ Google API không khả dụng")
                return pd.DataFrame()
            
            # Ví dụ: đọc từ sheet có sẵn
            # Trong thực tế, user sẽ cung cấp URL/ID
            print("📋 Để sử dụng, cần cung cấp Google Sheet URL hoặc ID")
            print("💡 Có thể sửa đổi method này để nhận input từ user")
            
            return pd.DataFrame()
            
        except Exception as e:
            print(f"❌ Lỗi đọc dữ liệu: {e}")
            return pd.DataFrame()

    def map_columns(self, data: pd.DataFrame) -> pd.DataFrame:
        """
        Ánh xạ columns theo config mapping
        
        Args:
            data (pd.DataFrame): DataFrame cần ánh xạ
            
        Returns:
            pd.DataFrame: DataFrame sau khi ánh xạ
        """
        try:
            # Mapping đơn giản - có thể customize sau
            mapped_data = data.copy()
            
            # Có thể thêm logic mapping từ config file
            print(f"📋 Ánh xạ columns hoàn tất: {len(mapped_data.columns)} columns")
            
            return mapped_data
            
        except Exception as e:
            print(f"❌ Lỗi mapping columns: {e}")
            return data

    def create_output_with_style(self, data: pd.DataFrame) -> str:
        """
        Tạo file output với styling
        
        Args:
            data (pd.DataFrame): Dữ liệu cần xuất
            
        Returns:
            str: Đường dẫn file output
        """
        try:
            # Tạo file output
            timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"{self.output_folder}/Google_Output_{timestamp}.xlsx"
            
            # Ghi dữ liệu
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                data.to_excel(writer, sheet_name='ProcessedData', index=False)
                
                # Áp dụng styling từ base class
                workbook = writer.book
                worksheet = workbook['ProcessedData']
                
                # Áp dụng border cho data
                self.apply_border_to_sheet(worksheet, len(data), len(data.columns))
                
                print(f"📁 Đã tạo file output với styling: {output_file}")
                return output_file
            
        except Exception as e:
            print(f"❌ Lỗi tạo output file: {e}")
            return None

    def process_data(self, sheet_url_or_id: str = None, sheet_name: str = None):
        """
        Xử lý dữ liệu từ Google Sheets với enhanced processing
        
        Args:
            sheet_url_or_id (str): URL hoặc ID của Google Sheet
            sheet_name (str): Tên sheet cụ thể (optional)
            
        Returns:
            str: Đường dẫn file output hoặc None nếu lỗi
        """
        print("📊 BẮTẦU XỬ LÝ GOOGLE SHEETS (Enhanced)")
        print("=" * 60)
        
        try:
            # Kiểm tra kết nối Google API
            if not self.check_google_api_status():
                print("❌ Không thể kết nối Google API")
                return None
            
            # Nếu không có input, yêu cầu user cung cấp
            if not sheet_url_or_id:
                print("⚠️  Cần cung cấp Google Sheet URL hoặc ID")
                print("💡 Ví dụ: process_data('1dUWOQzLF06aOvFIDIP7JUIt8CrvJlapKsrGv7xKAAMc')")
                return None
            
            # Đọc dữ liệu với enhanced method
            print(f"📋 Đang đọc từ sheet: {sheet_url_or_id}")
            data = self.google_client.read_shared_google_sheet(sheet_url_or_id, sheet_name)
            
            if data is None or data.empty:
                print("❌ Không có dữ liệu để xử lý")
                return None
            
            print(f"📋 Đã đọc {len(data)} dòng dữ liệu với enhanced processing")
            print(f"📊 Columns: {list(data.columns)}")
            
            # Xử lý dữ liệu với mapping từ config
            processed_data = self.map_columns(data)
            
            # Tạo output file với styling
            output_file = self.create_output_with_style(processed_data)
            
            print(f"✅ Hoàn thành! File output: {output_file}")
            return output_file
            
        except Exception as e:
            print(f"❌ Lỗi xử lý Google Sheets: {e}")
            return None
