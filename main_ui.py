"""
School Process UI Launcher - Fixed Version
Giao diện người dùng hiện đại cho ứng dụng School Process
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
from pathlib import Path
import time

# Thêm project root vào Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))


class SplashScreen:
    """Splash screen khi khởi động ứng dụng"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("School Process")
        self.root.geometry("400x300")
        self.root.resizable(False, False)
        
        # Remove title bar
        self.root.overrideredirect(True)
        
        # Center window
        self.center_window()
        
        # Configure background
        self.root.configure(bg='#1976D2')
        
        self.setup_ui()
        
    def center_window(self):
        """Căn giữa cửa sổ"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
    def setup_ui(self):
        """Thiết lập UI splash screen"""
        # Main frame
        main_frame = tk.Frame(self.root, bg='#1976D2')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Logo/Icon
        logo_label = tk.Label(main_frame, 
                            text="🏫",
                            font=('Segoe UI', 48),
                            bg='#1976D2',
                            fg='white')
        logo_label.pack(pady=(40, 20))
        
        # Title
        title_label = tk.Label(main_frame,
                             text="School Process",
                             font=('Segoe UI', 20, 'bold'),
                             bg='#1976D2',
                             fg='white')
        title_label.pack(pady=(0, 10))
        
        # Subtitle
        subtitle_label = tk.Label(main_frame,
                                text="Ứng dụng xử lý dữ liệu trường học",
                                font=('Segoe UI', 12),
                                bg='#1976D2',
                                fg='#E3F2FD')
        subtitle_label.pack(pady=(0, 30))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(main_frame,
                                     variable=self.progress_var,
                                     maximum=100,
                                     length=300)
        progress_bar.pack(pady=(0, 10))
        
        # Status label
        self.status_var = tk.StringVar(value="Đang khởi động...")
        status_label = tk.Label(main_frame,
                              textvariable=self.status_var,
                              font=('Segoe UI', 10),
                              bg='#1976D2',
                              fg='#E3F2FD')
        status_label.pack()
        
        # Version label
        version_label = tk.Label(main_frame,
                               text="Version 1.0.1 - UI Fixed",
                               font=('Segoe UI', 8),
                               bg='#1976D2',
                               fg='#BBDEFB')
        version_label.pack(side='bottom', pady=(20, 0))
        
    def update_progress(self, value, status=""):
        """Cập nhật tiến trình"""
        self.progress_var.set(value)
        if status:
            self.status_var.set(status)
        self.root.update()
        
    def close(self):
        """Đóng splash screen"""
        self.root.destroy()


def check_dependencies():
    """Kiểm tra dependencies"""
    required_modules = {
        'pandas': 'pandas',
        'requests': 'requests', 
        'google-auth': 'google.auth',
        'google-auth-oauthlib': 'google_auth_oauthlib',
        'google-auth-httplib2': 'google_auth_httplib2',
        'google-api-python-client': 'googleapiclient',
        'openpyxl': 'openpyxl'
    }
    
    missing_modules = []
    
    for package_name, import_name in required_modules.items():
        try:
            __import__(import_name)
            print(f"✅ {package_name} - OK")
        except ImportError as e:
            print(f"❌ {package_name} - Missing: {e}")
            missing_modules.append(package_name)
            
    return missing_modules


def show_dependency_error(missing_modules):
    """Hiển thị lỗi dependencies"""
    root = tk.Tk()
    root.withdraw()  # Hide main window
    
    error_msg = f"""
Thiếu các module cần thiết:

{chr(10).join(f"• {module}" for module in missing_modules)}

Vui lòng cài đặt bằng lệnh:
pip install {" ".join(missing_modules)}

Hoặc:
pip install -r requirements.txt
    """
    
    messagebox.showerror("Thiếu Dependencies", error_msg)
    root.destroy()


def load_app_with_splash():
    """Load app với splash screen"""
    print("🏫 School Process - Khởi động ứng dụng...")
    
    # Create splash screen
    splash = SplashScreen()
    print("✅ Đã tạo splash screen")
    
    try:
        # Step 1: Check dependencies
        splash.update_progress(20, "Đang kiểm tra dependencies...")
        print("📦 Kiểm tra dependencies...")
        
        missing = check_dependencies()
        if missing:
            print(f"❌ Thiếu dependencies: {missing}")
            splash.close()
            show_dependency_error(missing)
            return None
            
        # Step 2: Load config
        splash.update_progress(40, "Đang load cấu hình...")
        print("⚙️ Load cấu hình...")
        
        from config.config_manager import get_config
        config = get_config()
        print("✅ Đã load config")
        
        # Step 3: Import UI
        splash.update_progress(60, "Đang khởi tạo UI...")
        print("🖥️ Khởi tạo UI...")
        
        from ui.main_window import SchoolProcessMainWindow
        print("✅ Đã import UI module")
        
        # Step 4: Create app
        splash.update_progress(80, "Đang tạo giao diện...")
        print("🎨 Tạo main window...")
        
        app = SchoolProcessMainWindow()
        print("✅ Đã tạo main window")
        
        # Step 5: Complete
        splash.update_progress(100, "Hoàn tất!")
        print("🎉 Sẵn sàng hiển thị UI...")
        
        # Wait a moment
        time.sleep(0.5)
        
        # Close splash
        splash.close()
        print("✅ Splash screen đã đóng")
        
        return app
        
    except Exception as e:
        print(f"❌ Lỗi khởi động: {e}")
        import traceback
        traceback.print_exc()
        
        splash.close()
        
        # Show error
        root = tk.Tk()
        root.withdraw()
        error_msg = f"""
Lỗi khởi động ứng dụng: {str(e)}

Vui lòng kiểm tra:
1. Cấu hình trong file .env
2. Quyền truy cập file/thư mục
3. Kết nối mạng

Chi tiết lỗi:
{str(e)}
        """
        messagebox.showerror("Lỗi Khởi động", error_msg)
        root.destroy()
        return None


def main():
    """Entry point chính"""
    try:
        # Load app with splash screen
        app = load_app_with_splash()
        
        if app is None:
            print("❌ Không thể khởi động ứng dụng")
            return
            
        # Run main app
        print("🚀 Khởi động main app...")
        app.run()
        print("✅ App đã kết thúc")
        
    except Exception as e:
        print(f"❌ Lỗi nghiêm trọng: {e}")
        import traceback
        traceback.print_exc()
        
        # Fallback error dialog
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Lỗi nghiêm trọng", 
                               f"Không thể khởi động ứng dụng:\n{str(e)}")
            root.destroy()
        except:
            # Final fallback
            print("Không thể hiển thị dialog lỗi. Thoát ứng dụng.")
            input("Nhấn Enter để thoát...")

if __name__ == "__main__":
    main()
