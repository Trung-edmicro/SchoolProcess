"""
Google OAuth 2.0 Client for Drive API
Giải pháp thay thế Service Account để upload files lên Google Drive
Author: Assistant  
Date: 2025-07-27
"""

import os
import json
import pickle
from pathlib import Path
from typing import Optional

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
    import io
    GOOGLE_OAUTH_AVAILABLE = True
except ImportError:
    GOOGLE_OAUTH_AVAILABLE = False

from utils.menu_utils import print_status


class GoogleOAuthDriveClient:
    """Google OAuth 2.0 Drive Client với user storage quota"""
    
    # Drive API scopes
    SCOPES = [
        'https://www.googleapis.com/auth/drive.file',
        'https://www.googleapis.com/auth/drive'
    ]
    
    def __init__(self):
        """Khởi tạo OAuth Drive client"""
        if not GOOGLE_OAUTH_AVAILABLE:
            raise ImportError("Google OAuth libraries chưa được cài đặt. Chạy: pip install google-auth-oauthlib")
        
        self.credentials = None
        self.drive_service = None
        self.token_file = "config/oauth_token.pickle"
        self.credentials_file = "config/oauth_credentials.json"
        
        # Tạo thư mục config nếu chưa có
        os.makedirs("config", exist_ok=True)
        
        self._load_credentials()
        
        if self.credentials and self.credentials.valid:
            self._build_drive_service()
    
    def _load_credentials(self):
        """Load OAuth credentials từ file"""
        # Load từ token file nếu có
        if os.path.exists(self.token_file):
            try:
                with open(self.token_file, 'rb') as token:
                    self.credentials = pickle.load(token)
                print_status("✅ Đã load OAuth credentials từ cache", "success")
            except Exception as e:
                print_status(f"⚠️ Lỗi load token cache: {e}", "warning")
                self.credentials = None
        
        # Refresh token nếu hết hạn
        if self.credentials and self.credentials.expired and self.credentials.refresh_token:
            try:
                print_status("🔄 Đang refresh OAuth token...", "info")
                self.credentials.refresh(Request())
                self._save_credentials()
                print_status("✅ Đã refresh token thành công", "success")
            except Exception as e:
                print_status(f"❌ Lỗi refresh token: {e}", "error")
                self.credentials = None
        
        # Tạo credentials mới nếu cần
        if not self.credentials or not self.credentials.valid:
            self._create_new_credentials()
    
    def _create_new_credentials(self):
        """Tạo OAuth credentials mới thông qua browser flow"""
        if not os.path.exists(self.credentials_file):
            print_status(f"❌ Không tìm thấy OAuth credentials file: {self.credentials_file}", "error")
            print("\n💡 HƯỚNG DẪN SETUP OAUTH CREDENTIALS:")
            print("   1. Truy cập: https://console.cloud.google.com/")
            print("   2. Chọn project hoặc tạo project mới")
            print("   3. APIs & Services → Library → Enable 'Google Drive API'")
            print("   4. APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client IDs")
            print("   5. Application type: Desktop application")
            print("   6. Download JSON file và đổi tên thành 'oauth_credentials.json'")
            print(f"   7. Đặt file vào: {self.credentials_file}")
            print("\n🔗 Chi tiết: https://developers.google.com/drive/api/quickstart/python")
            return
        
        try:
            print_status("🔄 Đang tạo OAuth flow...", "info")
            flow = InstalledAppFlow.from_client_secrets_file(
                self.credentials_file, self.SCOPES)
            
            print("🌐 Mở trình duyệt để xác thực OAuth...")
            print("   👆 Vui lòng login với Google account và cấp quyền truy cập Drive")
            
            # Chạy OAuth flow
            self.credentials = flow.run_local_server(port=0)
            
            # Lưu credentials
            self._save_credentials()
            print_status("✅ OAuth setup thành công!", "success")
            
        except Exception as e:
            print_status(f"❌ Lỗi OAuth flow: {e}", "error")
            self.credentials = None
    
    def _save_credentials(self):
        """Lưu credentials vào file"""
        try:
            with open(self.token_file, 'wb') as token:
                pickle.dump(self.credentials, token)
            print_status(f"💾 Đã lưu OAuth token", "success")
        except Exception as e:
            print_status(f"⚠️ Lỗi lưu token: {e}", "warning")
    
    def _build_drive_service(self):
        """Tạo Google Drive service"""
        try:
            self.drive_service = build('drive', 'v3', credentials=self.credentials)
            print_status("✅ Đã kết nối Google Drive API với OAuth", "success")
            return True
        except Exception as e:
            print_status(f"❌ Lỗi tạo Drive service: {e}", "error")
            return False
    
    def is_authenticated(self) -> bool:
        """Kiểm tra xem đã xác thực thành công chưa"""
        return (self.credentials is not None and 
                self.credentials.valid and 
                self.drive_service is not None)
    
    def test_connection(self) -> bool:
        """Test kết nối Drive API"""
        if not self.is_authenticated():
            return False
        
        try:
            # Test bằng cách list 1 file
            results = self.drive_service.files().list(
                pageSize=1,
                fields="files(id, name)"
            ).execute()
            
            print_status("✅ OAuth Drive API test thành công", "success")
            return True
            
        except Exception as e:
            print_status(f"❌ OAuth Drive API test thất bại: {e}", "error")
            return False
    
    def get_user_info(self) -> dict:
        """Lấy thông tin user đã login OAuth"""
        if not self.is_authenticated():
            return {}
        
        try:
            about = self.drive_service.about().get(fields="user").execute()
            user = about.get('user', {})
            
            return {
                'email': user.get('emailAddress', 'N/A'),
                'name': user.get('displayName', 'N/A'),
                'photo': user.get('photoLink', 'N/A')
            }
        except Exception as e:
            print_status(f"⚠️ Lỗi lấy user info: {e}", "warning")
            return {}
    
    def upload_file_to_folder_id(self, local_path: str, folder_id: str, filename: str = None) -> Optional[str]:
        """
        Upload file lên Google Drive với folder ID cụ thể (backward compatibility)
        
        Args:
            local_path: Đường dẫn file local
            folder_id: ID folder đích trên Drive  
            filename: Tên file (mặc định lấy từ local_path)
            
        Returns:
            str: URL của file đã upload, None nếu thất bại
        """
        if not self.is_authenticated():
            print_status("❌ Chưa xác thực OAuth", "error")
            return None
        
        if not os.path.exists(local_path):
            print_status(f"❌ File không tồn tại: {local_path}", "error")
            return None
        
        if not filename:
            filename = os.path.basename(local_path)
        
        try:
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
                fields='id,webViewLink'
            ).execute()
            
            file_id = file.get('id')
            file_url = file.get('webViewLink')
            
            print_status(f"✅ OAuth upload thành công: {filename}", "success")
            print(f"   📁 Folder ID: {folder_id}")
            print(f"   🔗 URL: {file_url}")
            
            return file_url
            
        except Exception as e:
            print_status(f"❌ Lỗi OAuth upload {filename}: {e}", "error")
            return None

    def upload_file(self, local_path: str, folder_name: str = None, filename: str = None) -> Optional[str]:
        """
        Upload file lên Google Drive sử dụng OAuth (có user storage quota)
        
        Args:
            local_path: Đường dẫn file local
            folder_name: Tên folder đích (sẽ tự động tạo nếu chưa có, None = root)
            filename: Tên file (mặc định lấy từ local_path)
            
        Returns:
            str: URL của file đã upload, None nếu thất bại
        """
        if not self.is_authenticated():
            print_status("❌ Chưa xác thực OAuth", "error")
            return None
        
        if not os.path.exists(local_path):
            print_status(f"❌ File không tồn tại: {local_path}", "error")
            return None
        
        if not filename:
            filename = os.path.basename(local_path)
        
        try:
            # Tìm hoặc tạo folder
            folder_id = None
            if folder_name:
                folder_id = self._get_or_create_folder(folder_name)
                if not folder_id:
                    print_status(f"❌ Không thể tạo folder: {folder_name}", "error")
                    return None
            
            # Metadata cho file
            file_metadata = {
                'name': filename
            }
            
            # Thêm parent folder nếu có
            if folder_id:
                file_metadata['parents'] = [folder_id]
            
            # Upload file
            media = MediaFileUpload(local_path, resumable=True)
            file = self.drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id,webViewLink'
            ).execute()
            
            file_id = file.get('id')
            file_url = file.get('webViewLink')
            
            print_status(f"✅ OAuth upload thành công: {filename}", "success")
            print(f"   📁 Folder: {folder_name or 'Root'}")
            print(f"   🔗 URL: {file_url}")
            
            return file_url
            
        except Exception as e:
            print_status(f"❌ Lỗi OAuth upload {filename}: {e}", "error")
            return None
    
    def _get_or_create_folder(self, folder_name: str) -> Optional[str]:
        """Tìm hoặc tạo folder trên Drive"""
        try:
            # Tìm folder có sẵn
            query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder'"
            results = self.drive_service.files().list(
                q=query,
                spaces='drive',
                fields='files(id, name)'
            ).execute()
            
            folders = results.get('files', [])
            
            if folders:
                # Folder đã tồn tại
                folder_id = folders[0]['id']
                print_status(f"📁 Sử dụng folder có sẵn: {folder_name}", "info")
                return folder_id
            else:
                # Tạo folder mới
                print_status(f"📁 Tạo folder mới: {folder_name}", "info")
                folder_metadata = {
                    'name': folder_name,
                    'mimeType': 'application/vnd.google-apps.folder'
                }
                
                folder = self.drive_service.files().create(
                    body=folder_metadata,
                    fields='id'
                ).execute()
                
                return folder.get('id')
                
        except Exception as e:
            print_status(f"❌ Lỗi tạo folder {folder_name}: {e}", "error")
            return None
    
    def revoke_credentials(self):
        """Thu hồi OAuth credentials"""
        try:
            if self.credentials:
                self.credentials.revoke(Request())
            
            # Xóa token file
            if os.path.exists(self.token_file):
                os.remove(self.token_file)
                print_status("✅ Đã thu hồi OAuth credentials", "success")
            
            self.credentials = None
            self.drive_service = None
            
        except Exception as e:
            print_status(f"⚠️ Lỗi thu hồi credentials: {e}", "warning")
    
    def list_files_in_folder(self, folder_id: str):
        """Liệt kê tất cả file trong folder"""
        if not self.is_authenticated():
            return []
        
        try:
            query = f"'{folder_id}' in parents and trashed=false"
            results = self.drive_service.files().list(
                q=query,
                fields="nextPageToken, files(id, name, mimeType, modifiedTime, size)"
            ).execute()
            
            files = results.get('files', [])
            print_status(f"✅ Tìm thấy {len(files)} file trong folder", "info")
            
            return files
            
        except Exception as e:
            print_status(f"❌ Lỗi khi liệt kê file: {e}", "error")
            return []
    
    def find_import_files(self, folder_id: str):
        """Tìm tất cả file import (bắt đầu bằng 'import_') trong folder"""
        if not self.is_authenticated():
            return []
        
        try:
            # Tìm file với pattern "import_" và phần mở rộng .xlsx/.xls
            query = f"parents in '{folder_id}' and trashed=false and name contains 'import_' and (name contains '.xlsx' or name contains '.xls')"
            
            results = self.drive_service.files().list(
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
            
            print_status(f"✅ Tìm thấy {len(import_files)} file import", "info")
            return import_files
                
        except Exception as e:
            print_status(f"❌ Lỗi tìm file import: {e}", "error")
            return []
    
    def download_file(self, file_id: str, filename: str, local_dir: str = "data/temp"):
        """Tải file từ Drive về local"""
        if not self.is_authenticated():
            return None
        
        try:
            # Tạo thư mục nếu chưa có
            os.makedirs(local_dir, exist_ok=True)
            local_path = os.path.join(local_dir, filename)
            
            # Download file
            request = self.drive_service.files().get_media(fileId=file_id)
            
            with open(local_path, 'wb') as f:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
            
            print_status(f"✅ Đã tải file: {filename}", "success")
            return local_path
            
        except Exception as e:
            print_status(f"❌ Lỗi download file {filename}: {e}", "error")
            return None


def create_oauth_credentials_template():
    """Tạo template OAuth credentials để user dễ setup"""
    template = {
        "installed": {
            "client_id": "YOUR_CLIENT_ID.apps.googleusercontent.com",
            "project_id": "your-project-id", 
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_secret": "YOUR_CLIENT_SECRET",
            "redirect_uris": ["http://localhost"]
        }
    }
    
    template_file = "config/oauth_credentials_template.json"
    try:
        os.makedirs(os.path.dirname(template_file), exist_ok=True)
        with open(template_file, 'w', encoding='utf-8') as f:
            json.dump(template, f, ensure_ascii=False, indent=2)
        
        print_status(f"✅ Đã tạo template: {template_file}", "success")
        print("💡 Thay thế YOUR_CLIENT_ID và YOUR_CLIENT_SECRET từ Google Cloud Console")
        print("💡 Sau đó đổi tên thành 'oauth_credentials.json'")
        
    except Exception as e:
        print_status(f"❌ Lỗi tạo template: {e}", "error")
