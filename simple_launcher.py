"""
School Process Simple Launcher
Launcher đơn giản không dùng splash screen
"""

import tkinter as tk
from tkinter import messagebox
import sys
from pathlib import Path

# Thêm project root vào Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))


def check_dependencies():
    """Kiểm tra dependencies cơ bản"""
    try:
        import pandas
        import requests
        import google.auth
        import googleapiclient
        import openpyxl
        print("✅ Tất cả dependencies đã có")
        return True
    except ImportError as e:
        print(f"❌ Thiếu dependency: {e}")
        messagebox.showerror("Thiếu Dependencies", 
                           f"Thiếu module: {e}\n\nVui lòng chạy: pip install -r requirements.txt")
        return False


def main():
    """Entry point đơn giản"""
    try:
        print("🏫 School Process - Khởi động đơn giản...")
        
        # Check dependencies
        if not check_dependencies():
            return
            
        # Load config
        print("⚙️ Load config...")
        from config.config_manager import get_config
        config = get_config()
        
        # Import and run UI
        print("🖥️ Khởi tạo UI...")
        try:
            from ui.main_window import SchoolProcessMainWindow
            print("✅ Import main window thành công")
        except Exception as e:
            print(f"❌ Lỗi import main window: {e}")
            print("🔄 Sử dụng debug version...")
            from debug_main_window import SchoolProcessMainWindow
            
        print("🎨 Tạo ứng dụng...")
        app = SchoolProcessMainWindow()
        
        print("🚀 Chạy ứng dụng...")
        app.run()
        print("✅ Ứng dụng đã đóng")
        
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        import traceback
        traceback.print_exc()
        
        # Show error dialog
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Lỗi", f"Không thể khởi động ứng dụng:\n{str(e)}")
        root.destroy()


if __name__ == "__main__":
    main()
