"""
OnLuyen API Configuration
Cấu hình API endpoints cho hệ thống OnLuyen
Author: Assistant
Date: 2025-07-26
"""

from typing import Dict, Any, Optional
import requests
from dataclasses import dataclass
from urllib.parse import urljoin
import json
import os


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
        import urllib3
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
                            print(f"   ❌ Brotli decompression failed: {e}")
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
        print(f"   🔐 Headers: {list(headers.keys())}")
        
        try:
            response = requests.post(url, headers=headers, params=params, timeout=30)
            result = self._process_response(response)
            
            # Nếu thành công và có access_token mới, lưu vào file login
            if result["success"] and save_to_login_file and result.get("data"):
                self._update_login_file_with_new_token(result["data"], login_file_path, year)
            
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
            
            # Cập nhật tokens nếu có trong response
            if "access_token" in response_data:
                login_data["tokens"]["access_token"] = response_data["access_token"]
                print(f"✅ Updated access_token in login file")
            
            if "refresh_token" in response_data:
                login_data["tokens"]["refresh_token"] = response_data["refresh_token"]
                print(f"✅ Updated refresh_token in login file")
            
            if "expires_in" in response_data:
                login_data["tokens"]["expires_in"] = response_data["expires_in"]
                print(f"✅ Updated expires_in in login file")
            
            if "expires_at" in response_data:
                login_data["tokens"]["expires_at"] = response_data["expires_at"]
                print(f"✅ Updated expires_at in login file")
            
            # Thêm thông tin về việc thay đổi năm học
            if year:
                login_data["last_year_change"] = {
                    "year": year,
                    "timestamp": self._get_current_timestamp(),
                    "status": "success"
                }
                print(f"✅ Added year change info: {year}")
            
            # Lưu lại file
            with open(login_file_path, 'w', encoding='utf-8') as f:
                json.dump(login_data, f, indent=2, ensure_ascii=False)
            
            print(f"✅ Login file updated successfully: {login_file_path}")
            
        except Exception as e:
            print(f"❌ Error updating login file: {e}")
    
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
    
    def _get_current_timestamp(self) -> str:
        """
        Lấy timestamp hiện tại theo format của hệ thống
        
        Returns:
            str: Timestamp string
        """
        from datetime import datetime
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def load_token_from_login_file(self, login_file_path: str = None) -> bool:
        """
        Load access_token từ file login JSON
        
        Args:
            login_file_path (str, optional): Đường dẫn file login cụ thể.
                                           Nếu None, sẽ tìm file login gần nhất
            
        Returns:
            bool: True nếu load thành công, False nếu thất bại
        """
        try:
            # Tìm file login gần nhất nếu không chỉ định
            if not login_file_path:
                login_file_path = self._find_latest_login_file()
            
            if not login_file_path:
                print("❌ Không tìm thấy file login để load token")
                return False
            
            # Đọc file login
            with open(login_file_path, 'r', encoding='utf-8') as f:
                login_data = json.load(f)
            
            # Lấy access_token
            access_token = login_data.get("tokens", {}).get("access_token")
            if access_token:
                self.set_auth_token(access_token)
                print(f"✅ Access token loaded from: {login_file_path}")
                print(f"   Token: {access_token[:20]}...")
                return True
            else:
                print("❌ Không tìm thấy access_token trong file login")
                return False
                
        except Exception as e:
            print(f"❌ Error loading token from login file: {e}")
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
                # Thêm padding nếu cần
                padding = len(payload) % 4
                if padding:
                    payload += '=' * (4 - padding)
                
                import base64
                import json
                decoded_bytes = base64.b64decode(payload)
                decoded = json.loads(decoded_bytes.decode('utf-8'))
                
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
                            print(f"   ❌ Brotli decompression failed: {e}")
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
    
    print(f"\n🔧 Environment Variables Used:")
    print(f"   ONLUYEN_AUTH_BASE_URL = {os.getenv('ONLUYEN_AUTH_BASE_URL', 'default: https://auth.onluyen.vn')}")
    print(f"   ONLUYEN_SCHOOL_API_BASE_URL = {os.getenv('ONLUYEN_SCHOOL_API_BASE_URL', 'default: https://school-api.onluyen.vn')}")


if __name__ == "__main__":
    print_api_config_summary()
