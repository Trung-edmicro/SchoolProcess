"""
Main UI Window cho School Process - Version đơn giản để debug
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path

# Thêm project root vào Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from config.config_manager import get_config

class SchoolProcessMainWindow:
    def __init__(self):
        """Khởi tạo main window"""
        print("🔧 Tạo SchoolProcessMainWindow...")
        try:
            self.config = get_config()
            print("✅ Đã load config")
            
            self.setup_main_window()
            print("✅ Đã setup main window")
            
            self.setup_simple_ui()
            print("✅ Đã setup UI")
            
        except Exception as e:
            print(f"❌ Lỗi trong __init__: {e}")
            import traceback
            traceback.print_exc()
            raise
        
    def setup_main_window(self):
        """Thiết lập cửa sổ chính"""
        self.root = tk.Tk()
        self.root.title("School Process - Ứng dụng xử lý dữ liệu trường học")
        self.root.geometry("1000x700")
        self.root.minsize(800, 500)
        
        # Center window
        self.center_window()
        
    def center_window(self):
        """Căn giữa cửa sổ"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
    def setup_simple_ui(self):
        """Setup UI đơn giản"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Title
        title = ttk.Label(main_frame, text="School Process", 
                         font=('Segoe UI', 20, 'bold'))
        title.grid(row=0, column=0, pady=(0, 20))
        
        # Description
        desc = ttk.Label(main_frame, 
                        text="Ứng dụng xử lý dữ liệu trường học\nPhiên bản đơn giản để kiểm tra UI",
                        font=('Segoe UI', 12),
                        justify='center')
        desc.grid(row=1, column=0, pady=(0, 30))
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=2, column=0, pady=20)
        
        # Test button
        test_btn = ttk.Button(buttons_frame, text="Test Button", 
                             command=self.test_function)
        test_btn.pack(side='left', padx=10)
        
        # Close button
        close_btn = ttk.Button(buttons_frame, text="Đóng", 
                              command=self.root.quit)
        close_btn.pack(side='left', padx=10)
        
        # Status
        self.status_var = tk.StringVar(value="UI đã sẵn sàng")
        status_label = ttk.Label(main_frame, textvariable=self.status_var,
                                foreground='green')
        status_label.grid(row=3, column=0, pady=(20, 0))
        
    def test_function(self):
        """Test function"""
        self.status_var.set("Test button đã được click!")
        messagebox.showinfo("Test", "UI hoạt động bình thường!")
        
    def run(self):
        """Chạy ứng dụng"""
        print("🚀 Chạy main window...")
        try:
            self.root.mainloop()
            print("✅ Main window đã đóng")
        except Exception as e:
            print(f"❌ Lỗi trong run: {e}")
            import traceback
            traceback.print_exc()

def main():
    """Entry point cho UI"""
    try:
        app = SchoolProcessMainWindow()
        app.run()
    except Exception as e:
        print(f"❌ Lỗi main: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
