"""
UI Application Launcher
Khởi động ứng dụng với giao diện đồ họa
"""

import sys
import os
from pathlib import Path

# Thêm project root vào Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

def main():
    """Entry point cho UI application"""
    try:
        # Import UI main window
        from ui.main_window import SchoolProcessMainWindow
        
        # Tạo và chạy ứng dụng
        app = SchoolProcessMainWindow()
        app.run()
        
    except ImportError as e:
        print(f"Lỗi import module: {e}")
        print("Vui lòng cài đặt các dependencies cần thiết.")
        input("Nhấn Enter để thoát...")
        
    except Exception as e:
        print(f"Lỗi khởi động ứng dụng: {e}")
        input("Nhấn Enter để thoát...")


if __name__ == "__main__":
    main()
