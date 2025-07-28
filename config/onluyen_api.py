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
                token = result["data"].get("token") or result["data"].get("access_token")
                if token:
                    self.set_auth_token(token)
            
            return result
            
        except Exception as e:
            return {
                "success": False,
                "status_code": None,
                "data": None,
                "error": str(e)
            }
    
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
