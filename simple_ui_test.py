"""
Test UI đơn giản để kiểm tra vấn đề
"""

import tkinter as tk
from tkinter import ttk
import sys
import os
from pathlib import Path

# Thêm project root vào Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

class SimpleTestUI:
    def __init__(self):
        print("🔧 Tạo simple test UI...")
        self.root = tk.Tk()
        self.root.title("Simple Test")
        self.root.geometry("400x300")
        
        # Test label
        label = tk.Label(self.root, text="Test UI - Nếu thấy cửa sổ này thì UI hoạt động OK!", 
                        font=('Arial', 12), wraplength=350)
        label.pack(pady=50)
        
        # Close button
        btn = tk.Button(self.root, text="Đóng", command=self.root.quit)
        btn.pack(pady=20)
        
        print("✅ Simple UI đã tạo xong")
        
    def run(self):
        print("🚀 Chạy simple UI...")
        self.root.mainloop()
        print("✅ Simple UI đã đóng")

if __name__ == "__main__":
    try:
        app = SimpleTestUI()
        app.run()
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        import traceback
        traceback.print_exc()
