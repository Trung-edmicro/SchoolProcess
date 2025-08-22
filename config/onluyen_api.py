"""
OnLuyen API Configuration
Cấu hình API endpoints cho hệ thống OnLuyen
Author: Assistant
Date: 2025-07-26
"""
import requests
import json
import os
import base64
import urllib3
import time
import traceback
import glob
import unicodedata
import pandas as pd

from typing import Dict, Any, Optional
from dataclasses import dataclass
from urllib.parse import urljoin
from urllib.parse import quote

@dataclass
class APIEndpoint:
    """Định nghĩa một API endpoint"""
    name: str
    method: str
    url: str
    default_params: Dict[str, Any] = None
    description: str = ""
    
    def __post_init__(self):
        if self.default_params is None:
            self.default_params = {}


class OnLuyenAPIConfig:
    """Cấu hình API cho hệ thống OnLuyen"""
    
    @classmethod
    def get_auth_base_url(cls) -> str:
        """Lấy Auth Base URL từ environment"""
        return os.getenv('ONLUYEN_AUTH_BASE_URL')
    
    @classmethod  
    def get_school_api_base_url(cls) -> str:
        """Lấy School API Base URL từ environment"""
        return os.getenv('ONLUYEN_SCHOOL_API_BASE_URL')
    
    @classmethod
    def _build_endpoints(cls) -> Dict[str, 'APIEndpoint']:
        """Build endpoints với URLs từ environment"""
        auth_base = cls.get_auth_base_url()
        school_api_base = cls.get_school_api_base_url()
        
        return {
            "login": APIEndpoint(
                name="login",
                method="POST", 
                url=f"{auth_base}/api/account/login",
                description="Đăng nhập tài khoản"
            ),
            
            "change_school_year": APIEndpoint(
                name="change_school_year",
                method="GET",
                url=f"{auth_base}/api/account/change-school-year",
                default_params={"codeApp": "SCHOOL"},
                description="Thay đổi năm học"
            ),
            
            "list_teacher": APIEndpoint(
                name="list_teacher",
                method="GET",
                url=f"{school_api_base}/school/list-teacher/%20/1",
                default_params={"pageSize": 10},
                description="Lấy danh sách giáo viên"
            ),
            
            "list_student": APIEndpoint(
                name="list_student", 
                method="GET",
                url=f"{school_api_base}/school/list-student",
                default_params={"pageIndex": 1, "pageSize": 15},
                description="Lấy danh sách học sinh"
            ),

            "delete_teacher": APIEndpoint(
                name="delete_teacher",
                method="DELETE",
                url=f"{school_api_base}/manage-user/delete-user",
                description="Xóa giáo viên theo email"
            )
        }
    
    @classmethod
    def get_endpoints(cls) -> Dict[str, 'APIEndpoint']:
        """Lấy tất cả endpoints (lazy loading)"""
        return cls._build_endpoints()
    
    # Default request settings
    DEFAULT_TIMEOUT = 30
    DEFAULT_HEADERS = {
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
        "Accept-Encoding": "gzip, deflate, br",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Origin": "https://onluyen.vn",
        "Referer": "https://onluyen.vn/"
    }
    
    @classmethod
    def get_endpoint(cls, name: str) -> Optional[APIEndpoint]:
        """
        Lấy thông tin endpoint theo tên
        
        Args:
            name (str): Tên endpoint
            
        Returns:
            Optional[APIEndpoint]: Thông tin endpoint hoặc None
        """
        endpoints = cls.get_endpoints()
        return endpoints.get(name)
    
    @classmethod
    def get_all_endpoints(cls) -> Dict[str, APIEndpoint]:
        """
        Lấy tất cả endpoints
        
        Returns:
            Dict[str, APIEndpoint]: Dictionary chứa tất cả endpoints
        """
        return cls.get_endpoints().copy()
    
    @classmethod
    def print_endpoints_summary(cls):
        """In tóm tắt tất cả endpoints"""
        print("🌐 ONLUYEN API ENDPOINTS CONFIGURATION")
        print("=" * 60)
        
        endpoints = cls.get_endpoints()
        for name, endpoint in endpoints.items():
            print(f"\n📋 {endpoint.name.upper()}:")
            print(f"   🔧 Method: {endpoint.method}")
            print(f"   🔗 URL: {endpoint.url}")
            if endpoint.default_params:
                print(f"   📊 Default Params: {endpoint.default_params}")
            print(f"   📝 Description: {endpoint.description}")
    
    @classmethod
    def validate_endpoints(cls) -> Dict[str, bool]:
        """
        Validate tất cả endpoints (check URL format)
        
        Returns:
            Dict[str, bool]: Kết quả validation cho từng endpoint
        """
        results = {}
        endpoints = cls.get_endpoints()
        
        for name, endpoint in endpoints.items():
            try:
                # Simple URL validation
                if endpoint.url.startswith(('http://', 'https://')):
                    results[name] = True
                else:
                    results[name] = False
            except Exception:
                results[name] = False
        
        return results


class OnLuyenAPIClient:
    """Client để gọi OnLuyen APIs"""
    
    def __init__(self, session: requests.Session = None):
        """
        Khởi tạo API client
        
        Args:
            session (requests.Session, optional): Session để sử dụng
        """
        self.session = session or requests.Session()
        self.session.headers.update(OnLuyenAPIConfig.DEFAULT_HEADERS)
        # Tạm thời bỏ qua SSL verification cho testing
        self.session.verify = False
        # Tắt cảnh báo SSL
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        self.auth_token = None
        
        # Đảm bảo environment variables được load
        self._ensure_env_loaded()
    
    def _ensure_env_loaded(self):
        """Đảm bảo environment variables được load từ .env"""
        # Kiểm tra xem có cần load .env không
        if not os.getenv('ONLUYEN_AUTH_BASE_URL'):
            self._load_env_file()
    
    def _load_env_file(self):
        """Load environment variables từ .env file"""
        from pathlib import Path
        env_file = Path('.env')
        if env_file.exists():
            with open(env_file, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#') and '=' in line:
                        key, value = line.split('=', 1)
                        os.environ[key] = value
            print(f"✅ OnLuyenAPIClient: Loaded .env file")
        else:
            print(f"❌ OnLuyenAPIClient: .env file not found")
    
    def _log_request_debug(self, method: str, url: str, payload: dict, headers: dict):
        """Log request details cho debug"""
        print(f"\n🔍 DEBUG REQUEST:")
        print(f"   Method: {method}")
        print(f"   URL: {url}")
        print(f"   Headers:")
        for key, value in headers.items():
            print(f"     {key}: {value}")
        print(f"   Payload:")
        for key, value in payload.items():
            # Ẩn password
            if key == "password":
                value = "*" * len(str(value))
            print(f"     {key}: {value}")
        print("=" * 50)
        
    def set_auth_token(self, token: str):
        """
        Đặt auth token cho requests
        
        Args:
            token (str): Auth token
        """
        self.auth_token = token
        self.session.headers["Authorization"] = f"Bearer {token}"
    
    def login(self, username: str, password: str) -> Dict[str, Any]:
        """
        Thực hiện login
        
        Args:
            username (str): Tên đăng nhập
            password (str): Mật khẩu
            
        Returns:
            Dict[str, Any]: Kết quả login
        """
        endpoint = OnLuyenAPIConfig.get_endpoint("login")
        
        payload = {
            "codeApp": "SCHOOL",
            "password": password,
            "rememberMe": True,
            "userName": username
        }
        
        # Debug logging
        self._log_request_debug(
            method=endpoint.method,
            url=endpoint.url,
            payload=payload,
            headers=self.session.headers
        )
        
        try:
            response = self.session.request(
                method=endpoint.method,
                url=endpoint.url,
                json=payload,
                timeout=OnLuyenAPIConfig.DEFAULT_TIMEOUT
            )
            
            print(f"\n📡 RESPONSE DEBUG:")
            print(f"   Status Code: {response.status_code}")
            print(f"   Headers: {dict(response.headers)}")
            print(f"   Content Length: {len(response.content)} bytes")
            
            # Xử lý response content với decompression
            response_data = None
            if response.content:
                try:
                    # Kiểm tra content encoding
                    content_encoding = response.headers.get('content-encoding', '').lower()
                    
                    if content_encoding == 'br':
                        # Xử lý Brotli compression
                        try:
                            import brotli
                            decompressed_content = brotli.decompress(response.content)
                            response_text = decompressed_content.decode('utf-8')
                            print(f"   ✅ Brotli decompressed successfully")
                        except ImportError:
                            print(f"   ❌ Brotli library not available")
                            # Fallback: try to decode as-is
                            response_text = response.text
                        except Exception as e:
                            # print(f"   ❌ Brotli decompression failed: {e}")
                            response_text = response.text
                    elif content_encoding == 'gzip':
                        # Requests tự động xử lý gzip
                        response_text = response.text
                        print(f"   ✅ Gzip handled automatically")
                    else:
                        # Không có compression hoặc compression khác
                        response_text = response.text
                        print(f"   ✅ No compression or auto-handled")
                    
                    print(f"   Decoded Content: {response_text[:200]}...")
                    response_data = json.loads(response_text)
                    print(f"   Response JSON Keys: {list(response_data.keys()) if isinstance(response_data, dict) else 'Not a dict'}")
                        
                except json.JSONDecodeError as e:
                    print(f"   JSON Parse Error: {e}")
                    print(f"   Raw Content (first 200): {response.content[:200]}")
                    response_data = None
                except Exception as e:
                    print(f"   Content Processing Error: {e}")
                    print(f"   Raw Content (first 200): {response.content[:200]}")
                    response_data = None
            
            result = {
                "success": response.status_code == 200,
                "status_code": response.status_code,
                "data": response_data,
                "error": None
            }
            
            # Tự động set token nếu login thành công
            if result["success"] and result["data"]:
                token = result["data"].get("access_token")
                if token:
                    self.set_auth_token(token)
                    print(f"✅ Access token automatically set after login")
            
            return result
            
        except Exception as e:
            return {
                "success": False,
                "status_code": None,
                "data": None,
                "error": str(e)
            }
    
    def change_year_v2(self, year: int, save_to_login_file: bool = True, login_file_path: str = None) -> Dict[str, Any]:
        """
        Thay đổi năm học bằng endpoint chính xác từ browser headers
        CHỈ GỌI API change year, KHÔNG login lại
        
        Args:
            year (int): Năm học mới (ví dụ: 2024, 2025)
            save_to_login_file (bool): Có lưu access_token mới vào file login không
            login_file_path (str, optional): Đường dẫn file login JSON để update
            
        Returns:
            Dict[str, Any]: Kết quả thay đổi năm học
        """
        if not self.auth_token:
            return {
                "success": False,
                "error": "Chưa có access_token. Vui lòng đăng nhập trước khi thay đổi năm học",
                "status_code": None,
                "data": None
            }
        
        # Kiểm tra năm học hiện tại
        current_info = self.get_current_school_year_info()
        if current_info.get('success') and current_info.get('school_year') == year:
            print(f"ℹ️ Đã ở năm học {year}, không cần chuyển đổi")
            return {
                "success": True,
                "message": f"Đã ở năm học {year}",
                "status_code": 200,
                "data": {"access_token": self.auth_token, "year": year}
            }
        
        # Sử dụng endpoint chính xác từ browser headers
        url = f"https://oauth.onluyen.vn/api/account/change-school-year/{year}"
        params = {'codeApp': 'SCHOOL'}
        
        headers = {
            'Authorization': f'Bearer {self.auth_token}',
            'Content-Type': 'application/json',
            'Origin': 'https://school.onluyen.vn',
            'Referer': 'https://school.onluyen.vn/',
            'Accept': 'application/json, text/plain, */*',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        print(f"\n📅 Changing school year to: {year}")
        print(f"   📍 URL: {url}")
        print(f"   📋 Params: {params}")
        
        try:
            response = requests.get(url, headers=headers, params=params, timeout=30, verify=False)
            
            print(f"\n📡 RESPONSE DEBUG:")
            print(f"   Status Code: {response.status_code}")
            print(f"   Content Length: {len(response.content)} bytes")
            
            # Xử lý response content với decompression
            response_data = None
            if response.content:
                try:
                    # Kiểm tra content encoding
                    content_encoding = response.headers.get('content-encoding', '').lower()
                    
                    if content_encoding == 'br':
                        # Xử lý Brotli compression
                        try:
                            import brotli
                            decompressed_content = brotli.decompress(response.content)
                            response_text = decompressed_content.decode('utf-8')
                            print(f"   ✅ Brotli decompressed successfully")
                        except ImportError:
                            print(f"   ❌ Brotli library not available")
                            response_text = response.text
                        except Exception as e:
                            # print(f"   ❌ Brotli decompression failed: {e}")
                            response_text = response.text
                    else:
                        response_text = response.text
                        print(f"   ✅ Response decoded successfully")
                    
                    print(f"   Response Content: {response_text[:200]}...")
                    response_data = json.loads(response_text)
                    print(f"   Response JSON Keys: {list(response_data.keys()) if isinstance(response_data, dict) else 'Not a dict'}")
                        
                except json.JSONDecodeError as e:
                    print(f"   JSON Parse Error: {e}")
                    response_data = None
                except Exception as e:
                    print(f"   Content Processing Error: {e}")
                    response_data = None
            
            result = {
                "success": response.status_code == 200,
                "status_code": response.status_code,
                "data": response_data,
                "error": None
            }
            
            # Nếu thành công, cập nhật token mới và lưu vào file
            if result["success"] and response_data and save_to_login_file:
                new_access_token = response_data.get('access_token')
                if new_access_token:
                    # Cập nhật token trong client
                    self.set_auth_token(new_access_token)
                    print(f"✅ Đã cập nhật access_token mới cho năm {year}")
                    
                    # Lưu token mới vào file với cấu trúc multi-year
                    self._update_login_file_with_new_token(response_data, login_file_path, year)
                else:
                    print(f"⚠️ API thành công nhưng không có access_token mới")
            elif not result["success"]:
                print(f"❌ Change year thất bại: {response_data}")
            
            return result
            
        except Exception as e:
            return {
                "success": False,
                "error": f"Lỗi khi thay đổi năm học: {str(e)}",
                "status_code": None,
                "data": None
            }

    def change_year(self, year: int, save_to_login_file: bool = True, login_file_path: str = None) -> Dict[str, Any]:
        """
        Thay đổi năm học sử dụng access_token để xác thực
        
        Args:
            year (int): Năm học mới (ví dụ: 2024, 2025)
            save_to_login_file (bool): Có lưu access_token mới vào file login không
            login_file_path (str, optional): Đường dẫn file login JSON để update
            
        Returns:
            Dict[str, Any]: Kết quả thay đổi năm học
        """
        if not self.auth_token:
            return {
                "success": False,
                "error": "Chưa có access_token. Vui lòng đăng nhập trước khi thay đổi năm học",
                "status_code": None,
                "data": None
            }
        
        # Thử nhiều endpoint URLs và methods khác nhau
        possible_configs = [
            # Thử với GET method
            {"url": f"{self.get_school_api_base_url()}/api/account/change-school-year", "method": "GET"},
            {"url": f"{self.get_auth_base_url()}/api/account/change-school-year", "method": "GET"},
            {"url": f"{self.get_school_api_base_url()}/account/change-school-year", "method": "GET"},
            {"url": f"{self.get_auth_base_url()}/account/change-school-year", "method": "GET"},
            {"url": f"{self.get_school_api_base_url()}/api/change-school-year", "method": "GET"},
            {"url": f"{self.get_auth_base_url()}/api/change-school-year", "method": "GET"},
            # Thử với POST method
            {"url": f"{self.get_school_api_base_url()}/api/account/change-school-year", "method": "POST"},
            {"url": f"{self.get_auth_base_url()}/api/account/change-school-year", "method": "POST"},
            {"url": f"{self.get_school_api_base_url()}/account/change-school-year", "method": "POST"},
            {"url": f"{self.get_auth_base_url()}/account/change-school-year", "method": "POST"},
        ]
        
        print(f"\n📅 Trying to change school year to: {year}")
        
        for i, config in enumerate(possible_configs, 1):
            print(f"\n🔄 Attempt {i}/{len(possible_configs)}")
            
            # Tạo URL với year parameter
            url_with_year = f"{config['url']}/{year}"
            
            # Tạo endpoint tạm thời với URL đã có year
            temp_endpoint = APIEndpoint(
                name="change_school_year",
                method="GET",
                url=url_with_year,
                default_params={"codeApp": "SCHOOL"},
                description="Thay đổi năm học"
            )
            
            # Gọi API với codeApp parameter
            params = {"codeApp": "SCHOOL"}
            
            print(f"   📍 URL: {url_with_year}")
            print(f"   📋 Params: {params}")
            
            result = self._make_request(temp_endpoint, params=params)
            
            # Nếu không phải lỗi 404, return kết quả (dù thành công hay thất bại)
            if result.get("status_code") != 404:
                print(f"   ✅ Found working endpoint: {config['url']}")
                
                # Nếu thành công và có access_token mới, lưu vào file login
                if result["success"] and save_to_login_file and result.get("data"):
                    self._update_login_file_with_new_token(result["data"], login_file_path, year)
                
                return result
            else:
                print(f"   ❌ 404 - Endpoint not found")
        
        # Nếu tất cả endpoints đều trả về 404
        return {
            "success": False,
            "error": f"Không tìm thấy endpoint thay đổi năm học. Đã thử {len(possible_configs)} URLs khác nhau.",
            "status_code": 404,
            "data": None
        }
    
    def _update_login_file_with_new_token(self, response_data: Dict[str, Any], 
                                        login_file_path: str = None, year: int = None):
        """
        Cập nhật file login JSON với access_token mới sau khi thay đổi năm học
        
        Args:
            response_data (Dict): Response data từ API change_year
            login_file_path (str, optional): Đường dẫn file login cụ thể
            year (int, optional): Năm học đã thay đổi
        """
        try:
            # Tìm file login gần nhất nếu không chỉ định
            if not login_file_path:
                login_file_path = self._find_latest_login_file()
            
            if not login_file_path:
                print("❌ Không tìm thấy file login để cập nhật")
                return
            
            # Đọc file login hiện tại
            with open(login_file_path, 'r', encoding='utf-8') as f:
                login_data = json.load(f)
            
            # Đảm bảo cấu trúc tokens_by_year tồn tại
            if "tokens_by_year" not in login_data:
                login_data["tokens_by_year"] = {}
            
            # Đảm bảo có year để cập nhật
            if not year:
                year = login_data.get("current_school_year", 2025)
            
            year_str = str(year)
            
            # Tạo hoặc cập nhật token info cho năm học cụ thể
            if year_str not in login_data["tokens_by_year"]:
                login_data["tokens_by_year"][year_str] = {}
            
            year_token = login_data["tokens_by_year"][year_str]
            
            # Cập nhật tokens nếu có trong response
            updated_fields = []
            if "access_token" in response_data:
                year_token["access_token"] = response_data["access_token"]
                updated_fields.append("access_token")
            
            if "refresh_token" in response_data:
                year_token["refresh_token"] = response_data["refresh_token"]
                updated_fields.append("refresh_token")
            
            if "expires_in" in response_data:
                year_token["expires_in"] = response_data["expires_in"]
                updated_fields.append("expires_in")
            
            if "expires_at" in response_data:
                year_token["expires_at"] = response_data["expires_at"]
                updated_fields.append("expires_at")
            
            if "userId" in response_data:
                year_token["user_id"] = response_data["userId"]
                updated_fields.append("user_id")
            
            if "display_name" in response_data:
                year_token["display_name"] = response_data["display_name"]
                updated_fields.append("display_name")
            
            if "account" in response_data:
                year_token["account"] = response_data["account"]
                updated_fields.append("account")
            
            # Cập nhật timestamp
            year_token["last_updated"] = self._get_current_timestamp()
            updated_fields.append("last_updated")
            
            # Cập nhật current_school_year
            login_data["current_school_year"] = year
            
            # Thêm thông tin về việc thay đổi năm học
            login_data["last_year_change"] = {
                "year": year,
                "timestamp": self._get_current_timestamp(),
                "status": "success",
                "updated_fields": updated_fields
            }
            
            # Lưu lại file
            with open(login_file_path, 'w', encoding='utf-8') as f:
                json.dump(login_data, f, indent=2, ensure_ascii=False)
            
            print(f"✅ Login file updated successfully: {login_file_path}")
            print(f"📅 Updated token fields for year {year}: {', '.join(updated_fields)}")
            
        except Exception as e:
            print(f"❌ Error updating login file: {e}")
            import traceback
            print(f"   Traceback: {traceback.format_exc()}")
    
    def _find_latest_login_file(self) -> str:
        """
        Tìm file login gần nhất trong thư mục output
        
        Returns:
            str: Đường dẫn file login gần nhất hoặc None
        """
        try:
            import glob
            from pathlib import Path
            
            # Tìm tất cả file login trong thư mục output
            output_dir = Path("data/output")
            login_pattern = output_dir / "onluyen_login_*.json"
            login_files = glob.glob(str(login_pattern))
            
            if not login_files:
                return None
            
            # Sắp xếp theo thời gian modified và lấy file mới nhất
            latest_file = max(login_files, key=os.path.getmtime)
            print(f"📁 Found latest login file: {latest_file}")
            return latest_file
            
        except Exception as e:
            print(f"❌ Error finding login file: {e}")
            return None

    def _check_valid_tokens_for_years(self, admin_email: str) -> Dict[int, bool]:
        """
        Kiểm tra token hợp lệ cho tất cả các năm học
        
        Args:
            admin_email (str): Email admin
            
        Returns:
            Dict[int, bool]: Dictionary mapping năm học -> có token hợp lệ
        """
        valid_tokens = {}
        years = [2024, 2025]
        
        for year in years:
            # Backup current token
            current_token = self.auth_token
            
            try:
                if self.load_token_from_login_file(admin_email, year):
                    valid_tokens[year] = True
                    print(f"   ✅ Token hợp lệ cho năm {year}")
                else:
                    valid_tokens[year] = False
                    print(f"   ❌ Token không hợp lệ cho năm {year}")
            except Exception as e:
                valid_tokens[year] = False
                print(f"   ❌ Lỗi kiểm tra token năm {year}: {e}")
            
            # Restore original token
            if current_token:
                self.set_auth_token(current_token)
        
        return valid_tokens
    
    def _get_current_timestamp(self) -> str:
        """
        Lấy timestamp hiện tại theo format của hệ thống
        
        Returns:
            str: Timestamp string
        """
        from datetime import datetime
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def load_token_from_login_file(self, admin_email: str = None, school_year: int = 2025, login_file_path: str = None) -> bool:
        """
        Load access_token từ file login JSON với hỗ trợ multi-year
        
        Args:
            admin_email (str, optional): Email admin để tìm file login cụ thể
            school_year (int, optional): Năm học cần load token. Default: 2025
            login_file_path (str, optional): Đường dẫn file login cụ thể
            
        Returns:
            bool: True nếu load thành công, False nếu thất bại
        """
        try:
            # Tìm file login
            if not login_file_path:
                if admin_email:
                    # Tìm file login theo admin_email
                    filename = f"onluyen_login_{admin_email.replace('@', '_').replace('.', '_')}.json"
                    login_file_path = f"data/output/{filename}"
                    
                    if not os.path.exists(login_file_path):
                        print(f"❌ Không tìm thấy file login cho {admin_email}")
                        return False
                else:
                    # Tìm file login gần nhất
                    login_file_path = self._find_latest_login_file()
                    
                    if not login_file_path:
                        print("❌ Không tìm thấy file login để load token")
                        return False
            
            # Đọc file login
            with open(login_file_path, 'r', encoding='utf-8') as f:
                login_data = json.load(f)
            
            # Kiểm tra cấu trúc file (new format vs old format)
            tokens_by_year = login_data.get("tokens_by_year", {})
            
            if tokens_by_year:
                # New format: multi-year tokens
                year_token = tokens_by_year.get(str(school_year))
                if year_token:
                    access_token = year_token.get("access_token")
                    if access_token and self._is_token_valid(access_token):
                        self.set_auth_token(access_token)
                        print(f"✅ Access token loaded for year {school_year}: {login_file_path}")
                        print(f"   Token: {access_token[:20]}...")
                        return True
                    else:
                        print(f"❌ Token cho năm {school_year} đã hết hạn hoặc không hợp lệ")
                        return False
                else:
                    print(f"❌ Không tìm thấy token cho năm học {school_year}")
                    return False
            else:
                # Old format: fallback compatibility
                access_token = login_data.get("tokens", {}).get("access_token")
                if access_token and self._is_token_valid(access_token):
                    self.set_auth_token(access_token)
                    print(f"✅ Access token loaded (old format): {login_file_path}")
                    print(f"   Token: {access_token[:20]}...")
                    return True
                else:
                    print("❌ Token cũ đã hết hạn hoặc không hợp lệ")
                    return False
                
        except Exception as e:
            print(f"❌ Error loading token from login file: {e}")
            return False
    
    def _is_token_valid(self, access_token: str) -> bool:
        """
        Kiểm tra xem token có còn hợp lệ không
        
        Args:
            access_token (str): Token cần kiểm tra
            
        Returns:
            bool: True nếu token còn hợp lệ
        """
        try:
            # Decode JWT token để kiểm tra expiry
            import base64
            import json
            from datetime import datetime
            
            if not access_token or not isinstance(access_token, str):
                print(f"   Invalid token format: {type(access_token)}")
                return False
            
            parts = access_token.split('.')
            if len(parts) < 3:
                print(f"   Token doesn't have 3 parts: {len(parts)}")
                return False
            
            # Decode payload (part 1)
            payload = parts[1]
            
            # Đảm bảo padding đúng cho base64
            # Base64 string length phải chia hết cho 4
            missing_padding = len(payload) % 4
            if missing_padding:
                payload += '=' * (4 - missing_padding)
            
            try:
                decoded_bytes = base64.urlsafe_b64decode(payload)
                decoded = json.loads(decoded_bytes.decode('utf-8'))
            except Exception as decode_error:
                print(f"   Error decoding token payload: {decode_error}")
                return False
            
            # Kiểm tra expiry time
            exp = decoded.get('exp')
            if exp:
                try:
                    exp_time = datetime.fromtimestamp(exp)
                    now = datetime.now()
                    
                    if exp_time > now:
                        print(f"   Token expires at: {exp_time}")
                        return True
                    else:
                        print(f"   Token expired at: {exp_time}")
                        return False
                except Exception as exp_error:
                    print(f"   Error checking expiry: {exp_error}")
                    return False
            
            # Nếu không có exp field, kiểm tra other fields để đảm bảo token hợp lệ
            required_fields = ['Email', 'userId', 'codeApp']
            for field in required_fields:
                if field not in decoded:
                    print(f"   Missing required field: {field}")
                    return False
            
            print(f"   Token valid (no exp field found, but has required fields)")
            return True
            
        except Exception as e:
            print(f"   Error validating token: {e}")
            return False
    
    def get_current_school_year_info(self) -> Dict[str, Any]:
        """
        Lấy thông tin năm học hiện tại từ access_token
        
        Returns:
            Dict[str, Any]: Thông tin năm học và user
        """
        try:
            if not self.auth_token:
                # Thử load token từ file login
                if not self.load_token_from_login_file():
                    return {"success": False, "error": "Không có access token"}
            
            # Decode JWT token manually (chỉ lấy payload, không verify)
            parts = self.auth_token.split('.')
            if len(parts) >= 2:
                # Decode payload (part 1)
                payload = parts[1]
                # Đảm bảo padding đúng cho base64
                missing_padding = len(payload) % 4
                if missing_padding:
                    payload += '=' * (4 - missing_padding)
                
                try:
                    decoded_bytes = base64.urlsafe_b64decode(payload)
                    decoded = json.loads(decoded_bytes.decode('utf-8'))
                except Exception as decode_error:
                    return {"success": False, "error": f"Lỗi decode token payload: {decode_error}"}
                
                school_year = decoded.get('SchoolYear')
                display_name = decoded.get('DisplayName', '')
                email = decoded.get('Email', '')
                
                return {
                    "success": True,
                    "school_year": school_year,
                    "display_name": display_name,
                    "email": email,
                    "decoded_payload": decoded
                }
            else:
                return {"success": False, "error": "Invalid token format"}
                
        except Exception as e:
            return {"success": False, "error": f"Lỗi decode token: {str(e)}"}
    
    def print_current_school_year_info(self):
        """In thông tin năm học hiện tại"""
        info = self.get_current_school_year_info()
        
        if info["success"]:
            print(f"\n📅 THÔNG TIN NĂM HỌC HIỆN TẠI:")
            if info.get("school_year"):
                print(f"   📚 Năm học: {info['school_year']}")
            if info.get("display_name"):
                print(f"   👤 Tài khoản: {info['display_name']}")
            if info.get("email"):
                print(f"   📧 Email: {info['email']}")
        else:
            print(f"❌ Không thể lấy thông tin năm học: {info.get('error')}")

    def ensure_valid_token(self, admin_email: str, password: str, school_year: int = 2025) -> bool:
        """
        Đảm bảo có token hợp lệ cho năm học cụ thể
        - Chỉ login DUY NHẤT lần đầu nếu chưa có token hợp lệ cho BẤT KỲ năm nào
        - Nếu có token năm khác, chỉ cần gọi change_year để lấy token năm mới
        
        Args:
            admin_email (str): Email admin
            password (str): Password
            school_year (int): Năm học cần token
            
        Returns:
            bool: True nếu có token hợp lệ
        """
        try:
            print(f"\n🔍 Kiểm tra token cho năm học {school_year}...")
            
            # Bước 1: Kiểm tra token cho năm học hiện tại
            if self.load_token_from_login_file(admin_email, school_year):
                print(f"✅ Đã có token hợp lệ cho năm {school_year}")
                return True
            
            print(f"⚠️ Không có token hợp lệ cho năm {school_year}")
            
            # Bước 2: Kiểm tra tất cả token cho các năm khác
            print(f"🔍 Kiểm tra token cho các năm khác...")
            valid_tokens = self._check_valid_tokens_for_years(admin_email)
            
            # Tìm năm có token hợp lệ (không phải năm hiện tại)
            valid_other_year = None
            for year, is_valid in valid_tokens.items():
                if year != school_year and is_valid:
                    valid_other_year = year
                    break
            
            if valid_other_year:
                # Có token cho năm khác - chỉ cần change year
                print(f"✅ Tìm thấy token hợp lệ cho năm {valid_other_year}")
                print(f"🔄 Chuyển năm học từ {valid_other_year} sang {school_year}...")
                
                # Load token cho năm có sẵn
                if self.load_token_from_login_file(admin_email, valid_other_year):
                    change_result = self.change_year_v2(school_year)
                    
                    if change_result.get('success', False):
                        print(f"✅ Đã chuyển năm và lưu token cho năm {school_year}")
                        return True
                    else:
                        print(f"❌ Chuyển năm thất bại: {change_result.get('error', 'Unknown error')}")
                        print(f"🔄 Fallback: Thực hiện login mới...")
                        # Fallback to login if change year fails
                else:
                    print(f"❌ Không thể load token cho năm {valid_other_year}")
            
            # Bước 3: Không có token hợp lệ cho bất kỳ năm nào - cần login
            print(f"🔐 Thực hiện login DUY NHẤT cho tài khoản {admin_email}...")
            
            login_result = self.login(admin_email, password)
            if not login_result.get('success', False):
                print(f"❌ Login thất bại: {login_result.get('error', 'Unknown error')}")
                return False
            
            print(f"✅ Login thành công!")
            
            # Kiểm tra năm học sau login
            current_info = self.get_current_school_year_info()
            current_year_from_token = current_info.get('school_year') if current_info.get('success') else None
            
            if current_year_from_token and current_year_from_token != school_year:
                print(f"� Token login mặc định cho năm {current_year_from_token}")
                
                # Lưu token cho năm hiện tại trước
                self._save_multi_year_token(admin_email, password, login_result, current_year_from_token)
                print(f"✅ Đã lưu token cho năm {current_year_from_token}")
                
                if school_year != current_year_from_token:
                    print(f"🔄 Chuyển sang năm mục tiêu {school_year}...")
                    # Chuyển sang năm mục tiêu
                    change_result = self.change_year_v2(school_year)
                    if change_result.get('success', False):
                        print(f"✅ Đã chuyển và lưu token cho năm {school_year}")
                        return True
                    else:
                        print(f"❌ Chuyển năm thất bại sau login: {change_result.get('error')}")
                        print(f"ℹ️ Vẫn có thể sử dụng token cho năm {current_year_from_token}")
                        return False
                else:
                    return True
            else:
                # Token login đã đúng năm mục tiêu
                self._save_multi_year_token(admin_email, password, login_result, school_year)
                print(f"✅ Đã login và lưu token cho năm {school_year}")
                return True
            
        except Exception as e:
            print(f"❌ Lỗi ensure_valid_token: {e}")
            import traceback
            print(f"   Traceback: {traceback.format_exc()}")
            return False
    
    def _save_multi_year_token(self, admin_email: str, password: str, login_result: dict, school_year: int):
        """
        Lưu token với cấu trúc multi-year
        
        Args:
            admin_email (str): Email admin
            password (str): Password  
            login_result (dict): Kết quả login
            school_year (int): Năm học
        """
        try:
            from datetime import datetime
            import os
            
            # File name cố định theo admin_email
            filename = f"onluyen_login_{admin_email.replace('@', '_').replace('.', '_')}.json"
            filepath = f"data/output/{filename}"
            
            # Tạo thư mục output nếu chưa có
            os.makedirs("data/output", exist_ok=True)
            
            # Lấy data từ response
            response_data = login_result.get('data', {})
            
            # Tạo token info cho năm học hiện tại
            current_year_token = {
                'access_token': response_data.get('access_token'),
                'refresh_token': response_data.get('refresh_token'),
                'expires_in': response_data.get('expires_in'),
                'expires_at': response_data.get('expires_at'),
                'user_id': response_data.get('userId'),
                'display_name': response_data.get('display_name'),
                'account': response_data.get('account'),
                'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            # Load existing data nếu có
            existing_data = {}
            if os.path.exists(filepath):
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        existing_data = json.load(f)
                except Exception as e:
                    print(f"⚠️ Không thể đọc file existing: {e}")
                    existing_data = {}
            
            # Cấu trúc mới với multi-year support
            login_info = {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'school_name': existing_data.get('school_name', 'Unknown'),
                'admin_email': admin_email,
                'admin_password': password,
                'drive_link': existing_data.get('drive_link', ''),
                'login_status': 'success',
                'current_school_year': school_year,  # Năm học hiện tại
                'tokens_by_year': existing_data.get('tokens_by_year', {}),  # Giữ lại tokens cũ
                'last_login': {
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'school_year': school_year,
                    'status_code': login_result.get('status_code'),
                    'response_keys': list(response_data.keys()) if response_data else []
                }
            }
            
            # Cập nhật token cho năm học hiện tại
            login_info['tokens_by_year'][str(school_year)] = current_year_token
            
            # Ghi file
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(login_info, f, ensure_ascii=False, indent=2)
            
            print(f"✅ Đã lưu token vào: {filepath}")
            print(f"📅 Token cho năm học {school_year} đã được cập nhật")
            
        except Exception as e:
            print(f"Lỗi lưu multi-year token: {e}")

    def switch_school_year(self, target_year: int = None, save_to_login_file: bool = True, 
                         login_file_path: str = None) -> Dict[str, Any]:
        """
        Chuyển đổi năm học giữa 2024 và 2025
        
        Args:
            target_year (int, optional): Năm học muốn chuyển đến. 
                                       Nếu None, sẽ tự động chuyển đổi giữa 2024-2025
            save_to_login_file (bool): Có lưu access_token mới vào file login không
            login_file_path (str, optional): Đường dẫn file login JSON để update
            
        Returns:
            Dict[str, Any]: Kết quả thay đổi năm học
        """
        if target_year is None:
            # Tự động chuyển đổi - mặc định chuyển sang 2025
            target_year = 2025
            print(f"📅 Auto-switching to school year: {target_year}")
        
        if target_year not in [2024, 2025]:
            return {
                "success": False,
                "error": f"Năm học không hợp lệ: {target_year}. Chỉ hỗ trợ 2024 hoặc 2025",
                "status_code": None,
                "data": None
            }
        
        return self.change_year(target_year, save_to_login_file, login_file_path)

    def get_teachers(self, page_size: int = 10, **kwargs) -> Dict[str, Any]:
        """
        Lấy danh sách giáo viên
        
        Args:
            page_size (int): Số lượng records mỗi page
            **kwargs: Các parameters khác
            
        Returns:
            Dict[str, Any]: Kết quả API call
        """
        endpoint = OnLuyenAPIConfig.get_endpoint("list_teacher")
        
        params = endpoint.default_params.copy()
        params["pageSize"] = page_size
        params.update(kwargs)
        
        return self._make_request(endpoint, params=params)
    
    def get_students(self, page_index: int = 1, page_size: int = 15, **kwargs) -> Dict[str, Any]:
        """
        Lấy danh sách học sinh
        
        Args:
            page_index (int): Chỉ số trang
            page_size (int): Số lượng records mỗi page
            **kwargs: Các parameters khác
            
        Returns:
            Dict[str, Any]: Kết quả API call
        """
        endpoint = OnLuyenAPIConfig.get_endpoint("list_student")
        
        params = endpoint.default_params.copy()
        params["pageIndex"] = page_index
        params["pageSize"] = page_size
        params.update(kwargs)
        
        return self._make_request(endpoint, params=params)
    
    def _make_request(self, endpoint: APIEndpoint, params: Dict = None, 
                     json_data: Dict = None) -> Dict[str, Any]:
        """
        Thực hiện API request với hỗ trợ decompression
        
        Args:
            endpoint (APIEndpoint): Endpoint info
            params (Dict, optional): URL parameters
            json_data (Dict, optional): JSON payload
            
        Returns:
            Dict[str, Any]: Kết quả API call
        """
        try:
            print(f"\n🔍 API REQUEST DEBUG:")
            print(f"   Method: {endpoint.method}")
            print(f"   URL: {endpoint.url}")
            print(f"   Params: {params}")
            print(f"   Auth Token: {'Set' if self.auth_token else 'Not set'}")
            
            response = self.session.request(
                method=endpoint.method,
                url=endpoint.url,
                params=params,
                json=json_data,
                timeout=OnLuyenAPIConfig.DEFAULT_TIMEOUT
            )
            
            print(f"\n📡 API RESPONSE DEBUG:")
            print(f"   Status Code: {response.status_code}")
            print(f"   Headers: {dict(response.headers)}")
            print(f"   Content Length: {len(response.content)} bytes")
            
            # Xử lý decompression và response data
            response_data = None
            if response.content:
                try:
                    # Kiểm tra content encoding
                    content_encoding = response.headers.get('content-encoding', '').lower()
                    
                    if content_encoding == 'br':
                        # Xử lý Brotli compression
                        try:
                            import brotli
                            decompressed_content = brotli.decompress(response.content)
                            response_text = decompressed_content.decode('utf-8')
                            print(f"   ✅ Brotli decompressed successfully")
                        except ImportError:
                            print(f"   ❌ Brotli library not available, installing...")
                            # Fallback: try to decode as-is
                            response_text = response.text
                        except Exception as e:
                            # print(f"   ❌ Brotli decompression failed: {e}")
                            response_text = response.text
                    elif content_encoding == 'gzip':
                        # Requests tự động xử lý gzip
                        response_text = response.text
                        print(f"   ✅ Gzip handled automatically by requests")
                    else:
                        # Không có compression hoặc compression khác
                        response_text = response.text
                        print(f"   ✅ No compression or auto-handled")
                    
                    print(f"   Decoded Content: {response_text[:200]}...")
                    response_data = json.loads(response_text)
                    
                    if isinstance(response_data, dict):
                        print(f"   Response JSON Keys: {list(response_data.keys())}")
                    elif isinstance(response_data, list):
                        print(f"   Response is list with {len(response_data)} items")
                    
                except json.JSONDecodeError as e:
                    print(f"   JSON Parse Error: {e}")
                    print(f"   Raw Content (first 200): {response.content[:200]}")
                    response_data = None
                except Exception as e:
                    print(f"   Content Processing Error: {e}")
                    print(f"   Raw Content (first 200): {response.content[:200]}")
                    response_data = None
            
            return {
                "success": response.status_code == 200,
                "status_code": response.status_code,
                "data": response_data,
                "error": None,
                "response": response_text[:500] + "..." if 'response_text' in locals() else None,
                "endpoint": endpoint.name
            }
            
        except Exception as e:
            return {
                "success": False,
                "status_code": None,
                "data": None,
                "error": str(e),
                "endpoint": endpoint.name
            }
    
    def _get_auth_headers(self) -> Dict[str, str]:
        """
        Lấy headers với authorization token
        
        Returns:
            Dict[str, str]: Headers với authorization
        """
        headers = OnLuyenAPIConfig.DEFAULT_HEADERS.copy()
        if self.auth_token:
            headers["Authorization"] = f"Bearer {self.auth_token}"
        return headers
    
    def delete_teacher(self, teacher_account: str) -> Dict[str, Any]:
        """
        Xóa giáo viên theo tài khoản

        Args:
            teacher_account (str): Tài khoản của giáo viên cần xóa

        Returns:
            Dict[str, Any]: Kết quả xóa
        """
        try:
            endpoint = OnLuyenAPIConfig.get_endpoint("delete_teacher")
            if not endpoint or not endpoint.url:
                return {"success": False, "status_code": None, "data": None, "error": "Delete endpoint không được cấu hình"}

            quoted_account = quote(teacher_account, safe='')
            url = f"{endpoint.url.rstrip('/')}/{quoted_account}"

            response = self.session.request(
                method=endpoint.method or "DELETE",
                url=url,
                headers=self._get_auth_headers(),
                timeout=OnLuyenAPIConfig.DEFAULT_TIMEOUT
            )
            
            response_data = {}
            try:
                response_data = response.json()
            except:
                response_data = {"message": response.text}
            
            return {
                "success": response.status_code in [200, 204],
                "status_code": response.status_code,
                "data": response_data,
                "error": None if response.status_code in [200, 204] else response_data.get("message", "Unknown error")
            }
            
        except Exception as e:
            return {
                "success": False,
                "status_code": None,
                "data": None,
                "error": str(e)
            }
    
    def bulk_delete_teachers(self, admin_email: str, admin_password: str, 
                           school_year: int = 2025, delay_seconds: float = 0.5) -> Dict[str, Any]:
        """
        Xóa toàn bộ giáo viên trong trường
        
        Args:
            admin_email (str): Email admin
            admin_password (str): Mật khẩu admin
            school_year (int): Năm học (mặc định 2025)
            delay_seconds (float): Thời gian chờ giữa các request (giây)
            
        Returns:
            Dict[str, Any]: Kết quả xóa hàng loạt
        """
        results = {
            "success": False,
            "total_teachers": 0,
            "deleted_count": 0,
            "failed_count": 0,
            "errors": [],
            "deleted_teachers": [],
            "failed_teachers": []
        }
        
        try:
            # 1. Đảm bảo đăng nhập và token hợp lệ
            print("🔐 Đăng nhập và kiểm tra token...")
            login_success = self.ensure_valid_token(admin_email, admin_password, school_year)
            
            if not login_success:
                results["errors"].append("Không thể đăng nhập hoặc xác thực token")
                return results
            
            # 2. Lấy danh sách tất cả giáo viên
            print("📋 Lấy danh sách tất cả giáo viên...")
            teachers_result = self.get_teachers(page_size=1000)
            
            if not teachers_result.get("success", False):
                results["errors"].append(f"Không thể lấy danh sách giáo viên: {teachers_result.get('error')}")
                return results
            
            teachers_data = teachers_result.get("data", [])
            
            # Normalize response: extract list from various possible shapes
            def _extract_list_from_response(obj):
                # direct list
                if isinstance(obj, list):
                    return obj
                # JSON string -> try parse
                if isinstance(obj, str):
                    try:
                        parsed = json.loads(obj)
                        return _extract_list_from_response(parsed)
                    except Exception:
                        return None
                # dict patterns
                if isinstance(obj, dict):
                    # common keys that hold list
                    for key in ("data", "items", "rows", "results"):
                        val = obj.get(key)
                        if isinstance(val, list):
                            return val
                    # sometimes nested: obj['data'] is dict with 'data' list or similar
                    for val in obj.values():
                        if isinstance(val, list):
                            return val
                    return None
                return None

            teachers_list = _extract_list_from_response(teachers_result.get("data", teachers_result))
            if teachers_list is None:
                results["errors"].append(f"Dữ liệu giáo viên không đúng định dạng: {type(teachers_result.get('data'))}")
                return results

            # Use normalized list
            teachers_data = teachers_list
            
            results["total_teachers"] = len(teachers_data)
            
            if results["total_teachers"] == 0:
                print("ℹ️ Không có giáo viên nào để xóa")
                results["success"] = True
                return results
            
            print(f"📊 Tìm thấy {results['total_teachers']} giáo viên cần xóa")
            
            # 3. Xóa từng giáo viên
            for i, teacher in enumerate(teachers_data, 1):
                print(f"\n🔍 Xử lý giáo viên {teacher}:")

                if isinstance(teacher, dict) and "teacherInfo" in teacher and isinstance(teacher["teacherInfo"], dict):
                    teacher_obj = teacher["teacherInfo"]

                teacher_id = teacher_obj.get("userId") or teacher_obj.get("_id")
                teacher_name = teacher_obj.get("displayName", "Unknown")
                teacher_account = teacher_obj.get("userName", "Unknown")
                
                if not teacher_id:
                    error_msg = f"Giáo viên {teacher_name} ({teacher_account}) không có ID"
                    print(f"❌ {error_msg}")
                    results["errors"].append(error_msg)
                    results["failed_count"] += 1
                    results["failed_teachers"].append({
                        "name": teacher_name,
                        "account": teacher_account,
                        "error": "Missing ID"
                    })
                    continue
                
                print(f"🗑️ [{i}/{results['total_teachers']}] Xóa {teacher_name} ({teacher_account})...")
                
                # Xóa giáo viên
                delete_result = self.delete_teacher(teacher_account)
                
                if delete_result.get("success", False):
                    print(f"✅ Đã xóa {teacher_name}")
                    results["deleted_count"] += 1
                    results["deleted_teachers"].append({
                        "id": teacher_id,
                        "name": teacher_name,
                        "account": teacher_account
                    })
                else:
                    error_msg = f"Lỗi xóa {teacher_name}: {delete_result.get('error')}"
                    print(f"❌ {error_msg}")
                    results["failed_count"] += 1
                    results["errors"].append(error_msg)
                    results["failed_teachers"].append({
                        "id": teacher_id,
                        "name": teacher_name,
                        "account": teacher_account,
                        "error": delete_result.get('error')
                    })
                
                # Chờ để tránh quá tải server
                if delay_seconds > 0 and i < results["total_teachers"]:
                    time.sleep(delay_seconds)
            
            # 4. Kiểm tra kết quả
            results["success"] = results["failed_count"] == 0
            
            print(f"\n📊 KẾT QUẢ XÓA HÀNG LOẠT:")
            print(f"   📚 Tổng số giáo viên: {results['total_teachers']}")
            print(f"   ✅ Đã xóa thành công: {results['deleted_count']}")
            print(f"   ❌ Thất bại: {results['failed_count']}")
            
            if results["failed_count"] > 0:
                print(f"\n❌ DANH SÁCH LỖI:")
                for error in results["errors"]:
                    print(f"   • {error}")
            
            return results
            
        except Exception as e:
            error_msg = f"Lỗi bulk delete: {str(e)}"
            print(f"❌ {error_msg}")
            results["errors"].append(error_msg)
            return results
    
    def selective_delete_teachers_from_excel(self, excel_file_path: str, 
                                           admin_email: str, admin_password: str,
                                           school_year: int = 2025, 
                                           delay_seconds: float = 0.5) -> Dict[str, Any]:
        """
        Xóa giáo viên dựa trên danh sách trong file Excel
        
        Args:
            excel_file_path (str): Đường dẫn file Excel chứa danh sách giáo viên cần xóa
            admin_email (str): Email admin
            admin_password (str): Mật khẩu admin
            school_year (int): Năm học (mặc định 2025)
            delay_seconds (float): Thời gian chờ giữa các request (giây)
            
        Returns:
            Dict[str, Any]: Kết quả xóa có chọn lọc
        """
        
        results = {
            "success": False,
            "excel_file": excel_file_path,
            "total_from_excel": 0,
            "matched_teachers": 0,
            "deleted_count": 0,
            "failed_count": 0,
            "not_found_count": 0,
            "errors": [],
            "deleted_teachers": [],
            "failed_teachers": [],
            "not_found_teachers": []
        }
        
        try:
            # 1. Kiểm tra file Excel
            if not os.path.exists(excel_file_path):
                results["errors"].append(f"File Excel không tồn tại: {excel_file_path}")
                return results
            
            # 2. Đọc danh sách giáo viên từ Excel
            print(f"📖 Đọc danh sách giáo viên từ: {excel_file_path}")
            try:
                df = pd.read_excel(excel_file_path)
                
                # Tìm cột chứa tài khoản (có thể là 'account', 'username', 'tài khoản', v.v.)
                account_column = None
                possible_columns = ['Tên đăng nhập','account', 'username', 'tài khoản', 'tai_khoan', 'email',]
                
                def _normalize(s):
                    s = str(s or '').strip().lower()
                    # loại bỏ dấu (diacritics) để so sánh ổn định
                    nfkd = unicodedata.normalize('NFKD', s)
                    return ''.join(ch for ch in nfkd if not unicodedata.combining(ch))
                
                possible_norm = [_normalize(p) for p in possible_columns]
                
                for col in df.columns:
                    col_norm = _normalize(col)
                    if col_norm in possible_norm:
                        account_column = col
                        break
                
                if account_column is None:
                    results["errors"].append("Không tìm thấy cột tài khoản trong Excel. Cần có cột: account, username, tài khoản, tai_khoan, hoặc email")
                    return results
                
                # Lấy danh sách tài khoản cần xóa
                accounts_to_delete = df[account_column].dropna().astype(str).tolist()
                results["total_from_excel"] = len(accounts_to_delete)
                
                print(f"📊 Tìm thấy {results['total_from_excel']} tài khoản trong Excel")
                
            except Exception as e:
                results["errors"].append(f"Lỗi đọc file Excel: {str(e)}")
                return results
            
            # 3. Đảm bảo đăng nhập và token hợp lệ
            print("🔐 Đăng nhập và kiểm tra token...")
            login_success = self.ensure_valid_token(admin_email, admin_password, school_year)
            
            if not login_success:
                results["errors"].append("Không thể đăng nhập hoặc xác thực token")
                return results
            
            # 4. Lấy danh sách tất cả giáo viên từ hệ thống
            print("📋 Lấy danh sách tất cả giáo viên từ hệ thống...")
            teachers_result = self.get_teachers(page_size=1000)
            
            if not teachers_result.get("success", False):
                results["errors"].append(f"Không thể lấy danh sách giáo viên: {teachers_result.get('error')}")
                return results
            
            all_teachers = teachers_result.get("data", [])
            
            # Normalize response: extract list from various possible shapes
            def _extract_list_from_response(obj):
                # direct list
                if isinstance(obj, list):
                    return obj
                # JSON string -> try parse
                if isinstance(obj, str):
                    try:
                        parsed = json.loads(obj)
                        return _extract_list_from_response(parsed)
                    except Exception:
                        return None
                # dict patterns
                if isinstance(obj, dict):
                    # common keys that hold list
                    for key in ("data", "items", "rows", "results"):
                        val = obj.get(key)
                        if isinstance(val, list):
                            return val
                    # sometimes nested: obj['data'] is dict with 'data' list or similar
                    for val in obj.values():
                        if isinstance(val, list):
                            return val
                    return None
                return None

            teachers_list = _extract_list_from_response(teachers_result.get("data", teachers_result))
            if teachers_list is None:
                results["errors"].append(f"Dữ liệu giáo viên không đúng định dạng: {type(teachers_result.get('data'))}")
                return results

            # Use normalized list
            teachers_data = teachers_list
            
            # 5. Tìm matching giáo viên
            teachers_to_delete = []
            
            for account in accounts_to_delete:
                found_teacher = None
                
                for teacher in teachers_data:
                    if isinstance(teacher, dict) and "teacherInfo" in teacher and isinstance(teacher["teacherInfo"], dict):
                        teacher_obj = teacher["teacherInfo"]

                    teacher_account = teacher_obj.get("userName", "Unknown")

                    if teacher_account.lower() == account.lower():
                        found_teacher = teacher
                        break
                
                if found_teacher:
                    teachers_to_delete.append(found_teacher)
                else:
                    results["not_found_count"] += 1
                    results["not_found_teachers"].append(account)
            
            results["matched_teachers"] = len(teachers_to_delete)
            
            # print(f"🔍 Kết quả khớp:")
            # print(f"   ✅ Tìm thấy: {results['matched_teachers']}")
            # print(f"   ❓ Không tìm thấy: {results['not_found_count']}")
            
            if results["not_found_count"] > 0:
                print(f"   📝 Danh sách không tìm thấy: {results['not_found_teachers']}")
            
            # 6. Xóa từng giáo viên đã khớp
            if results["matched_teachers"] > 0:
                print(f"\n🗑️ Bắt đầu xóa {results['matched_teachers']} giáo viên...")
                
                for i, teacher in enumerate(teachers_to_delete, 1):

                    if isinstance(teacher, dict) and "teacherInfo" in teacher and isinstance(teacher["teacherInfo"], dict):
                        teacher_obj = teacher["teacherInfo"]

                    teacher_id = teacher_obj.get("userId") or teacher_obj.get("_id")
                    teacher_name = teacher_obj.get("displayName", "Unknown")
                    teacher_account = teacher_obj.get("userName", "Unknown")
                    
                    if not teacher_id:
                        error_msg = f"Giáo viên {teacher_name} ({teacher_account}) không có ID"
                        print(f"❌ {error_msg}")
                        results["errors"].append(error_msg)
                        results["failed_count"] += 1
                        results["failed_teachers"].append({
                            "name": teacher_name,
                            "account": teacher_account,
                            "error": "Missing ID"
                        })
                        continue
                    
                    print(f"🗑️ [{i}/{results['matched_teachers']}] Xóa {teacher_name} ({teacher_account})...")
                    
                    # Xóa giáo viên
                    delete_result = self.delete_teacher(teacher_account)
                    
                    if delete_result.get("success", False):
                        print(f"✅ Đã xóa {teacher_name}")
                        results["deleted_count"] += 1
                        results["deleted_teachers"].append({
                            "id": teacher_id,
                            "name": teacher_name,
                            "account": teacher_account
                        })
                    else:
                        error_msg = f"Lỗi xóa {teacher_name}: {delete_result.get('error')}"
                        print(f"❌ {error_msg}")
                        results["failed_count"] += 1
                        results["errors"].append(error_msg)
                        results["failed_teachers"].append({
                            "id": teacher_id,
                            "name": teacher_name,
                            "account": teacher_account,
                            "error": delete_result.get('error')
                        })

                    # Chờ để tránh quá tải server
                    if delay_seconds > 0 and i < results["matched_teachers"]:
                        time.sleep(delay_seconds)

            # 7. Kiểm tra kết quả
            results["success"] = results["failed_count"] == 0
            
            # print(f"\n📊 KẾT QUẢ XÓA CÓ CHỌN LỌC:")
            # print(f"   📂 File Excel: {excel_file_path}")
            # print(f"   📋 Tổng số tài khoản trong Excel: {results['total_from_excel']}")
            # print(f"   🔍 Tìm thấy trong hệ thống: {results['matched_teachers']}")
            # print(f"   ❓ Không tìm thấy: {results['not_found_count']}")
            # print(f"   ✅ Đã xóa thành công: {results['deleted_count']}")
            # print(f"   ❌ Thất bại: {results['failed_count']}")
            
            if results["failed_count"] > 0:
                print(f"\n❌ DANH SÁCH LỖI:")
                for error in results["errors"]:
                    print(f"   • {error}")
            
            return results
            
        except Exception as e:
            error_msg = f"Lỗi selective delete: {str(e)}"
            print(f"❌ {error_msg}")
            results["errors"].append(error_msg)
            return results
    
    def test_connectivity(self) -> Dict[str, Any]:
        """
        Test kết nối đến các endpoints
        
        Returns:
            Dict[str, Any]: Kết quả test
        """
        results = {}
        
        for name, endpoint in OnLuyenAPIConfig.get_all_endpoints().items():
            if name == "login":
                # Skip login test vì cần credentials
                results[name] = {"status": "skipped", "reason": "requires_credentials"}
                continue
            
            try:
                response = self.session.head(
                    url=endpoint.url,
                    timeout=10
                )
                results[name] = {
                    "status": "success" if response.status_code < 500 else "error",
                    "status_code": response.status_code
                }
            except Exception as e:
                results[name] = {
                    "status": "error",
                    "error": str(e)
                }
        
        return results

def print_api_config_summary():
    """In tóm tắt cấu hình API"""
    print("🌐 ONLUYEN API CONFIGURATION SUMMARY")
    print("=" * 70)
    
    print(f"\n🏢 Base URLs (từ Environment):")
    print(f"   🔐 Auth: {OnLuyenAPIConfig.get_auth_base_url()}")
    print(f"   🏫 School API: {OnLuyenAPIConfig.get_school_api_base_url()}")
    
    print(f"\n⚙️ Default Settings:")
    print(f"   ⏱️  Timeout: {OnLuyenAPIConfig.DEFAULT_TIMEOUT}s")
    print(f"   📋 Headers: {OnLuyenAPIConfig.DEFAULT_HEADERS}")
    
    OnLuyenAPIConfig.print_endpoints_summary()
    
    print(f"\n✅ Validation Results:")
    validation_results = OnLuyenAPIConfig.validate_endpoints()
    for name, is_valid in validation_results.items():
        status = "✅ Valid" if is_valid else "❌ Invalid"
        print(f"   {name}: {status}")

if __name__ == "__main__":
    print_api_config_summary()
