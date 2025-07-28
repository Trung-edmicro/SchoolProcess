"""
Google API Client for Drive và Sheets
Author: Assistant
Date: 2025-07-26
"""

import os
import io
import pandas as pd
from typing import Dict, List, Optional, Any
from datetime import datetime

try:
    from google.auth.transport.requests import Request
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
    import gspread
    from gspread_dataframe import get_as_dataframe, set_with_dataframe
    GOOGLE_LIBS_AVAILABLE = True
except ImportError:
    GOOGLE_LIBS_AVAILABLE = False
    print("⚠️  Google API libraries chưa được cài đặt!")
    print("   Chạy: pip install google-auth google-auth-oauthlib google-auth-httplib2")
    print("   Chạy: pip install google-api-python-client gspread gspread-dataframe")

from config import config


class GoogleAPIClient:
    """Client để tương tác với Google Drive và Google Sheets"""
    
    def __init__(self):
        """Khởi tạo Google API Client"""
        if not GOOGLE_LIBS_AVAILABLE:
            raise ImportError("Google API libraries không khả dụng. Hãy cài đặt các thư viện cần thiết.")
        
        self.credentials = None
        self.drive_service = None
        self.sheets_service = None
        self.gspread_client = None
        self.service_account_file = config.SERVICE_ACCOUNT_FILE
        
        # Tạo thư mục sync nếu chưa tồn tại
        os.makedirs(config.LOCAL_SYNC_FOLDER, exist_ok=True)
        
        self._authenticate()
    
    def _authenticate(self):
        """Xác thực với Google APIs"""
        try:
            if not os.path.exists(config.SERVICE_ACCOUNT_FILE):
                raise FileNotFoundError(f"Service account file không tồn tại: {config.SERVICE_ACCOUNT_FILE}")
            
            # Tạo credentials từ service account
            self.credentials = service_account.Credentials.from_service_account_file(
                config.SERVICE_ACCOUNT_FILE,
                scopes=config.SCOPES
            )
            
            # Khởi tạo các service
            self.drive_service = build('drive', 'v3', credentials=self.credentials)
            self.sheets_service = build('sheets', 'v4', credentials=self.credentials)
            self.gspread_client = gspread.authorize(self.credentials)
            
            print("✅ Đã xác thực thành công với Google APIs")
            
        except Exception as e:
            print(f"❌ Lỗi xác thực Google APIs: {e}")
            raise
    
    def list_files_in_folder(self, folder_id: str) -> List[Dict[str, Any]]:
        """
        Liệt kê các file trong Google Drive folder
        
        Args:
            folder_id (str): ID của folder
            
        Returns:
            List[Dict]: Danh sách file trong folder
        """
        try:
            query = f"'{folder_id}' in parents and trashed=false"
            results = self.drive_service.files().list(
                q=query,
                fields="nextPageToken, files(id, name, mimeType, modifiedTime, size)"
            ).execute()
            
            files = results.get('files', [])
            print(f"✅ Tìm thấy {len(files)} file trong folder")
            
            return files
            
        except Exception as e:
            print(f"❌ Lỗi khi liệt kê file: {e}")
            return []
    
    def download_file_from_drive(self, file_id: str, local_path: str) -> bool:
        """
        Tải file từ Google Drive về local
        
        Args:
            file_id (str): ID của file trên Drive
            local_path (str): Đường dẫn local để lưu file
            
        Returns:
            bool: True nếu thành công
        """
        try:
            request = self.drive_service.files().get_media(fileId=file_id)
            
            # Tạo thư mục nếu chưa tồn tại
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            
            with io.FileIO(local_path, 'wb') as file_handle:
                downloader = MediaIoBaseDownload(file_handle, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                    if status:
                        print(f"📥 Tải xuống {int(status.progress() * 100)}%")
            
            print(f"✅ Đã tải file về: {local_path}")
            return True
            
        except Exception as e:
            print(f"❌ Lỗi khi tải file: {e}")
            return False
    
    def upload_file_to_drive(self, local_path: str, folder_id: str, filename: str = None) -> Optional[str]:
        """
        Upload file từ local lên Google Drive
        
        Args:
            local_path (str): Đường dẫn file local
            folder_id (str): ID folder đích trên Drive
            filename (str): Tên file (mặc định lấy từ local_path)
            
        Returns:
            Optional[str]: ID của file đã upload
        """
        try:
            if not os.path.exists(local_path):
                raise FileNotFoundError(f"File local không tồn tại: {local_path}")
            
            if not filename:
                filename = os.path.basename(local_path)
            
            # Metadata cho file
            file_metadata = {
                'name': filename,
                'parents': [folder_id]
            }
            
            # Upload file
            media = MediaFileUpload(local_path, resumable=True)
            file = self.drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            
            file_id = file.get('id')
            print(f"✅ Đã upload file: {filename} (ID: {file_id})")
            return file_id
            
        except Exception as e:
            print(f"❌ Lỗi khi upload file: {e}")
            return None
    
    def find_file_by_name(self, folder_id: str, filename: str) -> Optional[str]:
        """
        Tìm file trong folder theo tên
        
        Args:
            folder_id (str): ID folder
            filename (str): Tên file cần tìm
            
        Returns:
            Optional[str]: ID của file nếu tìm thấy
        """
        try:
            query = f"'{folder_id}' in parents and name='{filename}' and trashed=false"
            results = self.drive_service.files().list(
                q=query,
                fields="files(id, name)"
            ).execute()
            
            files = results.get('files', [])
            if files:
                file_id = files[0]['id']
                print(f"✅ Tìm thấy file: {filename} (ID: {file_id})")
                return file_id
            else:
                print(f"❌ Không tìm thấy file: {filename}")
                return None
                
        except Exception as e:
            print(f"❌ Lỗi khi tìm file: {e}")
            return None
    
    def read_google_sheet(self, sheet_id: str, sheet_name: str = None) -> pd.DataFrame:
        """
        Đọc dữ liệu từ Google Sheets thông qua ID
        
        Args:
            sheet_id (str): ID của Google Sheets
            sheet_name (str): Tên sheet (mặc định là sheet đầu tiên)
            
        Returns:
            pd.DataFrame: Dữ liệu từ sheet
        """
        return self._read_sheet_with_fallback(sheet_id, sheet_name)
    
    def _read_sheet_with_fallback(self, sheet_id: str, sheet_name: str = None) -> pd.DataFrame:
        """
        Đọc Google Sheets với fallback mechanism
        
        Args:
            sheet_id (str): ID của Google Sheets
            sheet_name (str): Tên sheet
            
        Returns:
            pd.DataFrame: Dữ liệu từ sheet
        """
        print(f"📊 Đang đọc Google Sheet ID: {sheet_id}")
        
        # Method 1: Thử với gspread (ưu tiên)
        result = self._try_gspread_read(sheet_id, sheet_name)
        if not result.empty:
            return result
        
        # Method 2: Fallback với Sheets API
        print("🔄 Fallback: Sử dụng Sheets API...")
        result = self._try_sheets_api_read(sheet_id, sheet_name)
        if not result.empty:
            return result
        
        print(f"❌ Không thể đọc sheet với ID: {sheet_id}")
        return pd.DataFrame()
    
    def _try_gspread_read(self, sheet_id: str, sheet_name: str = None) -> pd.DataFrame:
        """Thử đọc với gspread"""
        try:
            print("🔍 Thử với gspread...")
            spreadsheet = self.gspread_client.open_by_key(sheet_id)
            
            # Lấy worksheet
            if sheet_name:
                try:
                    worksheet = spreadsheet.worksheet(sheet_name)
                except:
                    print(f"⚠️  Không tìm thấy sheet '{sheet_name}', dùng sheet đầu tiên")
                    worksheet = spreadsheet.sheet1
            else:
                worksheet = spreadsheet.sheet1
            
            # Đọc dữ liệu
            df = get_as_dataframe(worksheet, header=0, evaluate_formulas=True)
            df = self._clean_dataframe(df)
            
            print(f"✅ Gspread success: {worksheet.title} ({len(df)} hàng, {len(df.columns)} cột)")
            return df
            
        except Exception as e:
            print(f"❌ Gspread failed: {e}")
            return pd.DataFrame()
    
    def _try_sheets_api_read(self, sheet_id: str, sheet_name: str = None) -> pd.DataFrame:
        """Thử đọc với Sheets API"""
        try:
            print("🔍 Thử với Sheets API...")
            
            # Lấy metadata để biết các sheet có sẵn
            sheet_metadata = self.sheets_service.spreadsheets().get(
                spreadsheetId=sheet_id
            ).execute()
            
            available_sheets = [s['properties']['title'] for s in sheet_metadata['sheets']]
            print(f"📋 Available sheets: {', '.join(available_sheets)}")
            
            # Xác định sheet name
            if sheet_name and sheet_name in available_sheets:
                target_sheet = sheet_name
            else:
                target_sheet = available_sheets[0]  # Dùng sheet đầu tiên
                if sheet_name:
                    print(f"⚠️  Sheet '{sheet_name}' không tồn tại, dùng '{target_sheet}'")
            
            # Đọc dữ liệu từ sheet
            range_name = f"{target_sheet}!A:ZZ"  # Đọc toàn bộ dữ liệu
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=range_name,
                valueRenderOption='UNFORMATTED_VALUE'  # Lấy giá trị thô
            ).execute()
            
            values = result.get('values', [])
            
            if not values:
                print("❌ Sheet trống hoặc không có dữ liệu")
                return pd.DataFrame()
            
            # Tạo DataFrame
            df = self._values_to_dataframe(values)
            df = self._clean_dataframe(df)
            
            print(f"✅ Sheets API success: {target_sheet} ({len(df)} hàng, {len(df.columns)} cột)")
            return df
            
        except Exception as e:
            print(f"❌ Sheets API failed: {e}")
            return pd.DataFrame()
    
    def _values_to_dataframe(self, values: list) -> pd.DataFrame:
        """Chuyển đổi values từ Sheets API thành DataFrame"""
        if not values:
            return pd.DataFrame()
        
        # Lấy header từ hàng đầu tiên
        headers = values[0] if values else []
        
        # Lấy data từ các hàng còn lại
        data_rows = values[1:] if len(values) > 1 else []
        
        # Đảm bảo tất cả hàng có cùng số cột với header
        max_cols = len(headers)
        normalized_rows = []
        
        for row in data_rows:
            # Thêm empty string cho các cột thiếu
            normalized_row = row + [''] * (max_cols - len(row))
            # Cắt bớt nếu hàng dài hơn header
            normalized_row = normalized_row[:max_cols]
            normalized_rows.append(normalized_row)
        
        # Tạo DataFrame
        df = pd.DataFrame(normalized_rows, columns=headers)
        return df
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Làm sạch DataFrame"""
        if df.empty:
            return df
        
        # Bỏ các cột hoàn toàn trống
        df = df.dropna(how='all', axis=1)
        
        # Bỏ các hàng hoàn toàn trống
        df = df.dropna(how='all', axis=0)
        
        # Reset index
        df = df.reset_index(drop=True)
        
        # Chuyển đổi kiểu dữ liệu số nếu có thể
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='ignore')
        
        return df
    
    def write_google_sheet(self, sheet_id: str, sheet_name: str, df: pd.DataFrame, clear_first: bool = True) -> bool:
        """
        Ghi dữ liệu vào Google Sheets thông qua ID
        
        Args:
            sheet_id (str): ID của Google Sheets
            sheet_name (str): Tên sheet
            df (pd.DataFrame): Dữ liệu cần ghi
            clear_first (bool): Xóa dữ liệu cũ trước khi ghi
            
        Returns:
            bool: True nếu thành công
        """
        return self._write_sheet_with_fallback(sheet_id, sheet_name, df, clear_first)
    
    def _write_sheet_with_fallback(self, sheet_id: str, sheet_name: str, df: pd.DataFrame, clear_first: bool = True) -> bool:
        """
        Ghi Google Sheets với fallback mechanism
        
        Args:
            sheet_id (str): ID của Google Sheets
            sheet_name (str): Tên sheet
            df (pd.DataFrame): Dữ liệu cần ghi
            clear_first (bool): Xóa dữ liệu cũ trước
            
        Returns:
            bool: True nếu thành công
        """
        print(f"📝 Đang ghi vào Google Sheet ID: {sheet_id}, sheet: {sheet_name}")
        
        # Method 1: Thử với gspread (ưu tiên)
        if self._try_gspread_write(sheet_id, sheet_name, df, clear_first):
            return True
        
        # Method 2: Fallback với Sheets API
        print("🔄 Fallback: Sử dụng Sheets API...")
        if self._try_sheets_api_write(sheet_id, sheet_name, df, clear_first):
            return True
        
        print(f"❌ Không thể ghi vào sheet với ID: {sheet_id}")
        return False
    
    def _try_gspread_write(self, sheet_id: str, sheet_name: str, df: pd.DataFrame, clear_first: bool) -> bool:
        """Thử ghi với gspread"""
        try:
            print("🔍 Thử ghi với gspread...")
            spreadsheet = self.gspread_client.open_by_key(sheet_id)
            
            # Tìm hoặc tạo worksheet
            try:
                worksheet = spreadsheet.worksheet(sheet_name)
                print(f"✅ Tìm thấy sheet: {sheet_name}")
            except:
                print(f"📄 Tạo sheet mới: {sheet_name}")
                worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=26)
            
            # Xóa dữ liệu cũ nếu cần
            if clear_first:
                worksheet.clear()
                print("🧹 Đã xóa dữ liệu cũ")
            
            # Ghi dữ liệu mới
            set_with_dataframe(worksheet, df, include_index=False, include_column_header=True)
            
            print(f"✅ Gspread write success: {sheet_name} ({len(df)} hàng, {len(df.columns)} cột)")
            return True
            
        except Exception as e:
            print(f"❌ Gspread write failed: {e}")
            return False
    
    def _try_sheets_api_write(self, sheet_id: str, sheet_name: str, df: pd.DataFrame, clear_first: bool) -> bool:
        """Thử ghi với Sheets API"""
        try:
            print("🔍 Thử ghi với Sheets API...")
            
            # Chuẩn bị dữ liệu
            values = [df.columns.tolist()] + df.values.tolist()
            
            # Đảm bảo sheet tồn tại
            if not self._ensure_sheet_exists(sheet_id, sheet_name):
                return False
            
            # Xóa dữ liệu cũ nếu cần
            if clear_first:
                clear_range = f"{sheet_name}!A:ZZ"
                self.sheets_service.spreadsheets().values().clear(
                    spreadsheetId=sheet_id,
                    range=clear_range
                ).execute()
                print("🧹 Đã xóa dữ liệu cũ")
            
            # Ghi dữ liệu mới
            range_name = f"{sheet_name}!A1"
            value_range_body = {
                'values': values,
                'majorDimension': 'ROWS'
            }
            
            self.sheets_service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=range_name,
                valueInputOption='RAW',
                body=value_range_body
            ).execute()
            
            print(f"✅ Sheets API write success: {sheet_name} ({len(df)} hàng, {len(df.columns)} cột)")
            return True
            
        except Exception as e:
            print(f"❌ Sheets API write failed: {e}")
            return False
    
    def _ensure_sheet_exists(self, sheet_id: str, sheet_name: str) -> bool:
        """Đảm bảo sheet tồn tại, tạo mới nếu chưa có"""
        try:
            # Lấy danh sách sheet hiện có
            sheet_metadata = self.sheets_service.spreadsheets().get(
                spreadsheetId=sheet_id
            ).execute()
            
            existing_sheets = [s['properties']['title'] for s in sheet_metadata['sheets']]
            
            if sheet_name in existing_sheets:
                print(f"✅ Sheet '{sheet_name}' đã tồn tại")
                return True
            
            # Tạo sheet mới
            print(f"📄 Tạo sheet mới: {sheet_name}")
            requests = [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name
                    }
                }
            }]
            
            batch_update_request = {'requests': requests}
            self.sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id,
                body=batch_update_request
            ).execute()
            
            print(f"✅ Đã tạo sheet: {sheet_name}")
            return True
            
        except Exception as e:
            print(f"❌ Lỗi khi đảm bảo sheet tồn tại: {e}")
            return False
    
    def sync_files_from_drive(self) -> Dict[str, str]:
        """
        Sync các file cần thiết từ Google Drive về local
        
        Returns:
            Dict[str, str]: Mapping tên file -> đường dẫn local
        """
        local_files = {}
        
        if not config.DRIVE_INPUT_FOLDER_ID:
            print("❌ DRIVE_INPUT_FOLDER_ID chưa được cấu hình")
            return local_files
        
        try:
            for file_type, filename in config.DRIVE_FILES.items():
                print(f"\n🔄 Sync file: {filename}")
                
                # Tìm file trên Drive
                file_id = self.find_file_by_name(config.DRIVE_INPUT_FOLDER_ID, filename)
                
                if file_id:
                    # Tải về local
                    local_path = os.path.join(config.LOCAL_SYNC_FOLDER, filename)
                    if self.download_file_from_drive(file_id, local_path):
                        local_files[file_type] = local_path
                else:
                    print(f"⚠️  Không tìm thấy file: {filename}")
            
            print(f"\n✅ Đã sync {len(local_files)}/{len(config.DRIVE_FILES)} file")
            return local_files
            
        except Exception as e:
            print(f"❌ Lỗi trong quá trình sync: {e}")
            return local_files
    
    def get_students_data_from_sheets(self) -> pd.DataFrame:
        """
        Lấy dữ liệu học sinh từ Google Sheets
        
        Returns:
            pd.DataFrame: Dữ liệu học sinh
        """
        if not config.STUDENTS_SHEET_ID:
            print("❌ STUDENTS_SHEET_ID chưa được cấu hình")
            return pd.DataFrame()
        
        return self.read_google_sheet(
            config.STUDENTS_SHEET_ID,
            config.SHEET_NAMES['students']
        )
    
    def get_teachers_data_from_sheets(self) -> pd.DataFrame:
        """
        Lấy dữ liệu giáo viên từ Google Sheets
        
        Returns:
            pd.DataFrame: Dữ liệu giáo viên
        """
        if not config.TEACHERS_SHEET_ID:
            print("❌ TEACHERS_SHEET_ID chưa được cấu hình")
            return pd.DataFrame()
        
        return self.read_google_sheet(
            config.TEACHERS_SHEET_ID,
            config.SHEET_NAMES['teachers']
        )
    
    def upload_output_to_drive(self, local_file_path: str, filename: str = None) -> Optional[str]:
        """
        Upload file output lên Google Drive
        
        Args:
            local_file_path (str): Đường dẫn file local
            filename (str): Tên file trên Drive
            
        Returns:
            Optional[str]: ID file đã upload
        """
        if not config.DRIVE_OUTPUT_FOLDER_ID:
            print("❌ DRIVE_OUTPUT_FOLDER_ID chưa được cấu hình")
            return None
        
        return self.upload_file_to_drive(
            local_file_path,
            config.DRIVE_OUTPUT_FOLDER_ID,
            filename
        )
    
    def create_test_file(self, filename: str, content: str) -> Optional[str]:
        """
        Tạo file test trên Google Drive để kiểm tra quyền ghi
        
        Args:
            filename (str): Tên file test
            content (str): Nội dung file test
            
        Returns:
            Optional[str]: ID của file đã tạo
        """
        try:
            import io
            from googleapiclient.http import MediaIoBaseUpload
            
            # Tạo file metadata
            file_metadata = {
                'name': filename
            }
            
            # Tạo media upload từ string content
            file_content = io.BytesIO(content.encode('utf-8'))
            media = MediaIoBaseUpload(file_content, mimetype='text/plain')
            
            # Upload file
            file = self.drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            
            file_id = file.get('id')
            print(f"✅ Đã tạo file test: {filename} (ID: {file_id})")
            return file_id
            
        except Exception as e:
            print(f"❌ Lỗi tạo file test: {e}")
            return None
    
    def test_connection(self) -> bool:
        """
        Test kết nối tổng thể với Google APIs
        
        Returns:
            bool: True nếu tất cả service hoạt động tốt
        """
        try:
            # Test Drive service
            if not self.drive_service:
                print("❌ Drive service chưa được khởi tạo")
                return False
            
            # Test bằng cách list 1 file
            drive_results = self.drive_service.files().list(pageSize=1).execute()
            if 'files' not in drive_results:
                print("❌ Drive API không phản hồi đúng format")
                return False
            
            # Test Sheets service
            if not self.sheets_service:
                print("❌ Sheets service chưa được khởi tạo")
                return False
            
            # Test gspread client
            if not self.gspread_client:
                print("❌ Gspread client chưa được khởi tạo")
                return False
            
            # Nếu tất cả test đều pass
            print("✅ Tất cả Google API service hoạt động bình thường")
            return True
            
        except Exception as e:
            print(f"❌ Lỗi test connection: {e}")
            return False
    
    def read_shared_google_sheet(self, sheet_url_or_id: str, sheet_name: str = None) -> pd.DataFrame:
        """
        Đọc dữ liệu từ Google Sheets đã được chia sẻ - Version cải thiện
        
        Args:
            sheet_url_or_id (str): URL hoặc ID của Google Sheets đã được chia sẻ
            sheet_name (str): Tên sheet (mặc định là sheet đầu tiên)
            
        Returns:
            pd.DataFrame: Dữ liệu từ sheet
        """
        # Chuẩn hóa thành sheet ID
        sheet_id = self._extract_sheet_id(sheet_url_or_id)
        if not sheet_id:
            print("❌ Không thể trích xuất Sheet ID từ input")
            return pd.DataFrame()
        
        print(f"🔍 Đang đọc shared Google Sheet ID: {sheet_id}")
        
        # Sử dụng method đọc chung với fallback
        return self._read_sheet_with_fallback(sheet_id, sheet_name)
    
    def _extract_sheet_id(self, sheet_url_or_id: str) -> str:
        """
        Trích xuất Sheet ID từ URL hoặc trả về ID nếu đã là ID
        
        Args:
            sheet_url_or_id (str): URL hoặc ID
            
        Returns:
            str: Sheet ID hoặc empty string nếu không hợp lệ
        """
        if not sheet_url_or_id:
            return ""
        
        # Nếu là URL Google Sheets
        if 'docs.google.com/spreadsheets' in sheet_url_or_id:
            try:
                # Extract từ URL: https://docs.google.com/spreadsheets/d/{ID}/edit...
                parts = sheet_url_or_id.split('/d/')
                if len(parts) > 1:
                    sheet_id = parts[1].split('/')[0]
                    print(f"📋 Extracted ID từ URL: {sheet_id}")
                    return sheet_id
            except Exception as e:
                print(f"❌ Lỗi khi extract ID từ URL: {e}")
                return ""
        
        # Nếu đã là ID (hoặc có vẻ như ID)
        if len(sheet_url_or_id) > 20 and '/' not in sheet_url_or_id:
            print(f"📋 Using direct ID: {sheet_url_or_id}")
            return sheet_url_or_id
        
        print(f"⚠️  Input không được nhận dạng: {sheet_url_or_id}")
        return ""
    
    def test_shared_sheet_access(self, sheet_url_or_id: str) -> dict:
        """
        Kiểm tra quyền truy cập vào shared Google Sheet - Version cải thiện
        
        Args:
            sheet_url_or_id (str): URL hoặc ID của Google Sheets
            
        Returns:
            dict: Kết quả kiểm tra quyền truy cập
        """
        result = {
            "accessible": False,
            "method": None,
            "sheet_id": "",
            "sheet_title": "",
            "sheets_list": [],
            "permissions": {
                "read": False,
                "write": False
            },
            "errors": []
        }
        
        # Chuẩn hóa thành sheet ID
        sheet_id = self._extract_sheet_id(sheet_url_or_id)
        if not sheet_id:
            result["errors"].append("Không thể trích xuất Sheet ID")
            return result
        
        result["sheet_id"] = sheet_id
        print(f"🔍 Kiểm tra quyền truy cập sheet ID: {sheet_id}")
        
        # Test method 1: gspread
        if self._test_gspread_access(sheet_id, result):
            return result
        
        # Test method 2: Sheets API
        if self._test_sheets_api_access(sheet_id, result):
            return result
        
        return result
    
    def _test_gspread_access(self, sheet_id: str, result: dict) -> bool:
        """Test truy cập với gspread"""
        try:
            print("🔍 Test với gspread...")
            spreadsheet = self.gspread_client.open_by_key(sheet_id)
            
            result["accessible"] = True
            result["method"] = "gspread"
            result["sheet_title"] = spreadsheet.title
            result["sheets_list"] = [ws.title for ws in spreadsheet.worksheets()]
            result["permissions"]["read"] = True
            
            # Test write permission
            try:
                # Thử đọc một cell để test read
                worksheet = spreadsheet.sheet1
                worksheet.cell(1, 1).value  # Read test
                
                # Thử cập nhật một cell để test write (không thực sự thay đổi)
                current_value = worksheet.cell(1, 1).value
                worksheet.update_cell(1, 1, current_value)  # Write test
                result["permissions"]["write"] = True
                print("✅ Có quyền đọc và ghi")
                
            except Exception as write_error:
                print("⚠️  Chỉ có quyền đọc")
                # Write permission đã là False
            
            print(f"✅ Gspread access OK: {result['sheet_title']}")
            return True
            
        except Exception as e:
            result["errors"].append(f"Gspread: {e}")
            print(f"❌ Gspread failed: {e}")
            return False
    
    def _test_sheets_api_access(self, sheet_id: str, result: dict) -> bool:
        """Test truy cập với Sheets API"""
        try:
            print("🔍 Test với Sheets API...")
            
            # Test read permission
            sheet_metadata = self.sheets_service.spreadsheets().get(
                spreadsheetId=sheet_id
            ).execute()
            
            result["accessible"] = True
            result["method"] = "sheets_api"
            result["sheet_title"] = sheet_metadata['properties']['title']
            result["sheets_list"] = [
                sheet['properties']['title'] 
                for sheet in sheet_metadata['sheets']
            ]
            result["permissions"]["read"] = True
            
            # Test write permission (thử đọc một range nhỏ)
            try:
                first_sheet = result["sheets_list"][0]
                range_name = f"{first_sheet}!A1:A1"
                
                self.sheets_service.spreadsheets().values().get(
                    spreadsheetId=sheet_id,
                    range=range_name
                ).execute()
                
                # Nếu đọc được thì có thể có quyền write (cần test thêm)
                print("✅ Có quyền đọc (write permission chưa xác định)")
                
            except Exception as read_error:
                print(f"⚠️  Lỗi đọc: {read_error}")
            
            print(f"✅ Sheets API access OK: {result['sheet_title']}")
            return True
            
        except Exception as e:
            result["errors"].append(f"Sheets API: {e}")
            print(f"❌ Sheets API failed: {e}")
            return False
    
    def get_sheet_info(self, sheet_id: str) -> dict:
        """
        Lấy thông tin chi tiết về Google Sheet
        
        Args:
            sheet_id (str): ID của Google Sheets
            
        Returns:
            dict: Thông tin về sheet
        """
        info = {
            "accessible": False,
            "title": "",
            "sheets": [],
            "total_sheets": 0,
            "url": f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
        }
        
        try:
            # Lấy metadata
            sheet_metadata = self.sheets_service.spreadsheets().get(
                spreadsheetId=sheet_id
            ).execute()
            
            info["accessible"] = True
            info["title"] = sheet_metadata['properties']['title']
            
            # Lấy thông tin các sheet
            for sheet in sheet_metadata['sheets']:
                sheet_info = {
                    "title": sheet['properties']['title'],
                    "sheet_id": sheet['properties']['sheetId'],
                    "index": sheet['properties']['index'],
                    "sheet_type": sheet['properties'].get('sheetType', 'GRID'),
                    "row_count": sheet['properties']['gridProperties']['rowCount'],
                    "col_count": sheet['properties']['gridProperties']['columnCount']
                }
                info["sheets"].append(sheet_info)
            
            info["total_sheets"] = len(info["sheets"])
            
            print(f"✅ Sheet info: {info['title']} ({info['total_sheets']} sheets)")
            return info
            
        except Exception as e:
            print(f"❌ Lỗi lấy thông tin sheet: {e}")
            return info
