"""
Environment Configuration Manager
Quản lý cấu hình môi trường từ file .env
Author: Assistant
Date: 2025-07-26
"""

import os
from typing import Optional, Dict, Any
from pathlib import Path


class ConfigManager:
    """Quản lý cấu hình từ file .env"""
    
    def __init__(self, env_file: str = ".env"):
        """
        Khởi tạo ConfigManager
        
        Args:
            env_file (str): Đường dẫn đến file .env
        """
        self.env_file = env_file
        self.config = {}
        self._load_env_file()
    
    def _load_env_file(self) -> None:
        """Đọc và parse file .env"""
        try:
            env_path = Path(self.env_file)
            if not env_path.exists():
                print(f"⚠️  File .env không tồn tại: {self.env_file}")
                return
            
            with open(env_path, 'r', encoding='utf-8') as f:
                for line_num, line in enumerate(f, 1):
                    line = line.strip()
                    
                    # Bỏ qua comment và dòng trống
                    if not line or line.startswith('#'):
                        continue
                    
                    # Parse key=value
                    if '=' in line:
                        key, value = line.split('=', 1)
                        key = key.strip()
                        value = value.strip()
                        
                        # Xử lý boolean values
                        if value.lower() in ['true', 'false']:
                            value = value.lower() == 'true'
                        # Xử lý số
                        elif value.isdigit():
                            value = int(value)
                        
                        self.config[key] = value
                        # Cũng set vào os.environ để các module khác sử dụng được
                        os.environ[key] = str(value)
            
            print(f"✅ Đã load {len(self.config)} config từ {self.env_file}")
            
        except Exception as e:
            print(f"❌ Lỗi đọc file .env: {e}")
    
    def get(self, key: str, default: Any = None) -> Any:
        """
        Lấy giá trị config
        
        Args:
            key (str): Key cần lấy
            default (Any): Giá trị mặc định nếu không tìm thấy
            
        Returns:
            Any: Giá trị config hoặc default
        """
        # Ưu tiên environment variable
        env_value = os.getenv(key)
        if env_value is not None:
            return env_value
        
        # Sau đó từ file .env
        return self.config.get(key, default)
    
    def get_google_config(self) -> Dict[str, Any]:
        """
        Lấy tất cả cấu hình Google API
        
        Returns:
            Dict[str, Any]: Dictionary chứa config Google API
        """
        return {
            'service_account_email': self.get('GOOGLE_SERVICE_ACCOUNT_EMAIL'),
            'test_sheet_id': self.get('GOOGLE_TEST_SHEET_ID'),
            'credentials_file': self.get('GOOGLE_CREDENTIALS_FILE', 'config/service_account.json'),
            'api_timeout': self.get('GOOGLE_API_TIMEOUT', 30),
            'max_retries': self.get('MAX_API_RETRIES', 3)
        }
    
    def get_paths_config(self) -> Dict[str, str]:
        """
        Lấy cấu hình đường dẫn
        
        Returns:
            Dict[str, str]: Dictionary chứa các đường dẫn
        """
        return {
            'input_dir': self.get('DATA_INPUT_DIR', 'data/input'),
            'temp_dir': self.get('DATA_TEMP_DIR', 'data/temp'),
            'output_dir': self.get('DATA_OUTPUT_DIR', 'data/output'),
            'config_dir': self.get('CONFIG_DIR', 'config'),
            'log_file': self.get('LOG_FILE', 'logs/school_process.log')
        }
    
    def get_output_config(self) -> Dict[str, str]:
        """
        Lấy cấu hình output
        
        Returns:
            Dict[str, str]: Dictionary chứa config output
        """
        return {
            'timestamp_format': self.get('TIMESTAMP_FORMAT', '%Y%m%d_%H%M%S'),
            'file_prefix': self.get('OUTPUT_FILE_PREFIX', 'SchoolProcess_Output')
        }
    
    def is_debug_mode(self) -> bool:
        """
        Kiểm tra có ở debug mode không
        
        Returns:
            bool: True nếu debug mode
        """
        return self.get('DEBUG_MODE', False)
    
    def is_demo_mode(self) -> bool:
        """
        Kiểm tra có ở demo mode không
        
        Returns:
            bool: True nếu demo mode
        """
        return self.get('DEMO_MODE', True)
    
    def get_demo_sheet_name(self) -> Optional[str]:
        """
        Lấy tên sheet demo
        
        Returns:
            Optional[str]: Tên sheet demo hoặc None
        """
        sheet_name = self.get('DEMO_SHEET_NAME', '')
        return sheet_name if sheet_name else None
    
    def get_onluyen_config(self) -> Dict[str, Any]:
        """
        Lấy cấu hình OnLuyen API
        
        Returns:
            Dict[str, Any]: Dictionary chứa config OnLuyen API
        """
        return {
            'auth_base_url': self.get('ONLUYEN_AUTH_BASE_URL', 'https://auth.onluyen.vn'),
            'school_api_base_url': self.get('ONLUYEN_SCHOOL_API_BASE_URL', 'https://school-api.onluyen.vn'),
            'username': self.get('ONLUYEN_USERNAME', ''),
            'password': self.get('ONLUYEN_PASSWORD', ''),
            'api_timeout': int(self.get('ONLUYEN_API_TIMEOUT', '30')),
            'default_page_size': int(self.get('ONLUYEN_DEFAULT_PAGE_SIZE', '10')),
            'student_page_size': int(self.get('ONLUYEN_STUDENT_PAGE_SIZE', '15')),
            'auto_login': str(self.get('ONLUYEN_AUTO_LOGIN', 'false')).lower() == 'true',
            'cache_token': str(self.get('ONLUYEN_CACHE_TOKEN', 'true')).lower() == 'true'
        }
    
    def print_config_summary(self) -> None:
        """In tóm tắt cấu hình"""
        print("\n📋 CẤU HÌNH HỆ THỐNG:")
        print("=" * 50)
        
        google_config = self.get_google_config()
        paths_config = self.get_paths_config()
        onluyen_config = self.get_onluyen_config()
        
        print(f"🌐 Google API:")
        print(f"   📧 Service Account: {google_config['service_account_email']}")
        print(f"   📄 Test Sheet ID: {google_config['test_sheet_id']}")
        print(f"   📁 Credentials: {google_config['credentials_file']}")
        
        print(f"\n🏢 OnLuyen API:")
        print(f"   🔐 Auth URL: {onluyen_config['auth_base_url']}")
        print(f"   🏫 School API: {onluyen_config['school_api_base_url']}")
        username_display = onluyen_config['username'] if onluyen_config['username'] else 'Not configured'
        print(f"   👤 Username: {username_display}")
        print(f"   ⏱️  Timeout: {onluyen_config['api_timeout']}s")
        print(f"   📊 Page Size: {onluyen_config['default_page_size']}")
        
        print(f"\n📁 Đường dẫn:")
        print(f"   📥 Input: {paths_config['input_dir']}")
        print(f"   📤 Output: {paths_config['output_dir']}")
        print(f"   🔧 Config: {paths_config['config_dir']}")
        
        print(f"\n⚙️  Chế độ:")
        print(f"   🔧 Debug: {self.is_debug_mode()}")
        print(f"   🎯 Demo: {self.is_demo_mode()}")
        
        environment = self.get('ENVIRONMENT', 'development')
        print(f"   🌍 Environment: {environment}")


    def set_env_value(self, key: str, value: str) -> bool:
        """
        Cập nhật (hoặc thêm) key=value trong file .env và đồng bộ vào memory  os.environ.
        Trả về True nếu thành công.
        """
        try:
            env_path = Path(self.env_file)
            lines = []
            if env_path.exists():
                with env_path.open('r', encoding='utf-8') as f:
                    lines = f.readlines()

            key_prefix = f"{key}="
            found = False
            for idx, line in enumerate(lines):
                stripped = line.strip()
                if not stripped or stripped.startswith('#'):
                    continue
                if stripped.split('=', 1)[0].strip() == key:
                    lines[idx] = f"{key_prefix}{value}\n"
                    found = True
                    break

            if not found:
                if lines and not lines[-1].endswith('\n'):
                    lines[-1] = lines[-1] + '\n'
                lines.append(f"{key_prefix}{value}\n")

            # Write back atomically
            with env_path.open('w', encoding='utf-8') as f:
                f.writelines(lines)

            # Update runtime config and environment
            self.config[key] = value
            os.environ[key] = str(value)
            return True
        except Exception as e:
            print(f"❌ Không thể cập nhật .env: {e}")
            return False

    def set_sheets_id(self, new_sheet_id: str) -> bool:
        """
        Convenience helper: cập nhật GOOGLE_TEST_SHEET_ID trong .env
        """
        return self.set_env_value('GOOGLE_TEST_SHEET_ID', new_sheet_id)

    def reload(self) -> None:
        """
        Tải lại file .env vào bộ nhớ (gọi sau khi thay đổi .env từ ngoài).
        """
        self.config = {}
        self._load_env_file()

# Global config instance
config_manager = ConfigManager()


def get_config() -> ConfigManager:
    """
    Lấy instance ConfigManager global
    
    Returns:
        ConfigManager: Instance config manager
    """
    return config_manager
