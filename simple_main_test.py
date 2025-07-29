"""
Test Main Window đơn giản
"""

import tkinter as tk
from tkinter import ttk
import sys
import os
from pathlib import Path

# Thêm project root vào Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from config.config_manager import get_config

class SimpleMainWindow:
    def __init__(self):
        print("🔧 Tạo simple main window...")
        try:
            self.config = get_config()
            print("✅ Đã load config")
            
            self.root = tk.Tk()
            self.root.title("School Process - Simple Test")
            self.root.geometry("800x600")
            
            # Simple UI
            frame = tk.Frame(self.root)
            frame.pack(fill='both', expand=True, padx=20, pady=20)
            
            title = tk.Label(frame, text="School Process", 
                           font=('Arial', 16, 'bold'))
            title.pack(pady=10)
            
            msg = tk.Label(frame, text="UI hoạt động OK!\nMainWindow có thể tạo được.", 
                          font=('Arial', 12))
            msg.pack(pady=20)
            
            # Close button
            btn = tk.Button(frame, text="Đóng", command=self.root.quit,
                          font=('Arial', 10))
            btn.pack(pady=10)
            
            print("✅ Simple main window đã tạo xong")
            
        except Exception as e:
            print(f"❌ Lỗi trong constructor: {e}")
            import traceback
            traceback.print_exc()
            raise
        
    def run(self):
        print("🚀 Chạy simple main window...")
        try:
            self.root.mainloop()
            print("✅ Simple main window đã đóng")
        except Exception as e:
            print(f"❌ Lỗi trong run: {e}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    try:
        app = SimpleMainWindow()
        app.run()
    except Exception as e:
        print(f"❌ Lỗi tổng: {e}")
        import traceback
        traceback.print_exc()
