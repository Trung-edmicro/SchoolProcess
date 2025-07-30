"""
Demo test Tksheet - Google Sheets like interface
Chạy file này để test xem Tksheet có hoạt động không
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path

# Thêm project root vào Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

try:
    import tksheet
    TKSHEET_AVAILABLE = True
    print("✅ Tksheet đã được cài đặt")
except ImportError:
    TKSHEET_AVAILABLE = False
    print("❌ Tksheet chưa được cài đặt. Chạy: pip install tksheet")


class TksheetDemo:
    """Demo Tksheet với styling Google Sheets"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Tksheet Demo - Google Sheets Style")
        self.root.geometry("1000x600")
        self.root.configure(bg='white')
        
        self.sheet = None
        self.setup_ui()
        
    def setup_ui(self):
        """Thiết lập UI demo"""
        if not TKSHEET_AVAILABLE:
            self.show_error()
            return
            
        # Header
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill='x', padx=10, pady=10)
        
        title_label = ttk.Label(header_frame,
                               text="📊 Tksheet Demo - Google Sheets Style",
                               font=('Segoe UI', 16, 'bold'),
                               foreground='#1a73e8')
        title_label.pack(side='left')
        
        # Buttons
        btn_frame = ttk.Frame(header_frame)
        btn_frame.pack(side='right')
        
        ttk.Button(btn_frame, text="🔄 Refresh", command=self.refresh_data).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="➕ Add Row", command=self.add_row).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="💾 Save", command=self.save_data).pack(side='left', padx=5)
        
        # Separator
        ttk.Separator(self.root, orient='horizontal').pack(fill='x', pady=5)
        
        # Tksheet container
        sheet_frame = ttk.Frame(self.root)
        sheet_frame.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        
        # Create Tksheet
        self.create_sheet(sheet_frame)
        
        # Status bar
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill='x', side='bottom', padx=10, pady=(0, 10))
        
        ttk.Separator(status_frame, orient='horizontal').pack(fill='x', pady=(0, 5))
        
        self.status_label = ttk.Label(status_frame, 
                                     text="✅ Demo Tksheet sẵn sàng - Bạn có thể edit cells trực tiếp!",
                                     font=('Segoe UI', 9),
                                     foreground='#137333')
        self.status_label.pack(side='left')
        
    def show_error(self):
        """Hiển thị lỗi khi tksheet chưa cài đặt"""
        error_frame = ttk.Frame(self.root)
        error_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        ttk.Label(error_frame,
                 text="⚠️ Module 'tksheet' chưa được cài đặt",
                 font=('Segoe UI', 16, 'bold'),
                 foreground='red').pack(pady=20)
        
        ttk.Label(error_frame,
                 text="Vui lòng chạy lệnh sau để cài đặt:\npip install tksheet",
                 font=('Segoe UI', 12),
                 justify='center').pack(pady=10)
        
        ttk.Button(error_frame, 
                  text="Đóng",
                  command=self.root.destroy).pack(pady=20)
        
    def create_sheet(self, parent):
        """Tạo Tksheet với demo data"""
        # Create sheet
        self.sheet = tksheet.Sheet(
            parent,
            page_up_down_select_row=True,
            expand_sheet_if_paste_too_big=True,
            empty_horizontal=0,
            empty_vertical=0,
            show_horizontal_grid=True,
            show_vertical_grid=True
        )
        
        # Configure Google Sheets styling
        self.sheet.set_options(
            font=('Segoe UI', 11),
            header_font=('Segoe UI', 11, 'bold'),
            table_bg='white',
            table_fg='black',
            table_selected_cells_bg='#e8f0fe',
            table_selected_cells_fg='#1a73e8',
            table_selected_rows_bg='#e8f0fe',
            table_selected_rows_fg='#1a73e8',
            table_selected_columns_bg='#e8f0fe', 
            table_selected_columns_fg='#1a73e8',
            header_bg='#f8f9fa',
            header_fg='#202124',
            header_selected_cells_bg='#d2e3fc',
            header_selected_cells_fg='#1a73e8',
            index_bg='#f8f9fa',
            index_fg='#202124',
            index_selected_cells_bg='#d2e3fc',
            index_selected_cells_fg='#1a73e8',
            top_left_bg='#f8f9fa',
            top_left_fg='#202124',
            outline_thickness=1,
            outline_color='#dadce0',
            grid_color='#dadce0',
            header_border_fg='#dadce0',
            index_border_fg='#dadce0'
        )
        
        # Enable editing features
        self.sheet.enable_bindings([
            "single_select",
            "row_select", 
            "column_select",
            "column_width_resize",
            "row_height_resize",
            "double_click_column_resize",
            "right_click_popup_menu",
            "copy",
            "paste",
            "delete",
            "edit_cell"
        ])
        
        # Set headers
        headers = ["STT", "Tên Trường", "Admin", "Mật Khẩu", "Link Driver", "SL GV", "SL HS", "Ghi Chú"]
        self.sheet.headers(headers)
        
        # Demo data
        demo_data = [
            [1, "TRƯỜNG ABC", "admin@abc.edu.vn", "123456", "https://drive.google.com/abc", "50", "500", "Đã nhập"],
            [2, "TRƯỜNG XYZ", "admin@xyz.edu.vn", "password", "https://drive.google.com/xyz", "30", "300", "Chưa hoàn thành"],
            [3, "TRƯỜNG DEF", "admin@def.edu.vn", "admin123", "https://drive.google.com/def", "45", "450", "Đang xử lý"],
            [4, "TRƯỜNG GHI", "admin@ghi.edu.vn", "secure123", "https://drive.google.com/ghi", "60", "600", "Hoàn thành"],
            [5, "TRƯỜNG JKL", "admin@jkl.edu.vn", "pass2023", "https://drive.google.com/jkl", "35", "350", "Cần kiểm tra"]
        ]
        
        # Set data
        self.sheet.set_sheet_data(demo_data)
        
        # Set column widths
        column_widths = [60, 150, 180, 100, 200, 80, 80, 120]
        for i, width in enumerate(column_widths):
            self.sheet.column_width(column=i, width=width)
        
        # Pack sheet
        self.sheet.pack(fill='both', expand=True)
        
        # Bind events
        self.sheet.extra_bindings([
            ("cell_select", self.on_cell_select),
            ("begin_edit_cell", self.on_begin_edit),
            ("end_edit_cell", self.on_end_edit)
        ])
        
    def on_cell_select(self, event):
        """Xử lý khi select cell"""
        try:
            row = event.row
            col = event.column
            if row is not None and col is not None:
                self.status_label.config(text=f"📍 Selected: Row {row + 1}, Column {col + 1}")
        except:
            pass
            
    def on_begin_edit(self, event):
        """Xử lý khi bắt đầu edit"""
        try:
            row = event.row
            col = event.column
            if row is not None and col is not None:
                self.status_label.config(text=f"✏️ Editing: Row {row + 1}, Column {col + 1}")
        except:
            pass
            
    def on_end_edit(self, event):
        """Xử lý khi kết thúc edit"""
        try:
            row = event.row
            col = event.column
            if row is not None and col is not None:
                self.status_label.config(text=f"✅ Saved: Row {row + 1}, Column {col + 1}")
        except:
            pass
            
    def refresh_data(self):
        """Refresh data"""
        if self.sheet:
            messagebox.showinfo("Refresh", "Dữ liệu đã được làm mới!")
            self.status_label.config(text="🔄 Dữ liệu đã được refresh")
            
    def add_row(self):
        """Thêm row mới"""
        if self.sheet:
            current_data = self.sheet.get_sheet_data()
            new_stt = len(current_data) + 1
            new_row = [new_stt, "", "", "", "", "", "", ""]
            current_data.append(new_row)
            self.sheet.set_sheet_data(current_data)
            self.status_label.config(text=f"➕ Đã thêm row {new_stt}")
            
    def save_data(self):
        """Lưu data"""
        if self.sheet:
            data = self.sheet.get_sheet_data()
            messagebox.showinfo("Save", f"Đã lưu {len(data)} rows!")
            self.status_label.config(text=f"💾 Đã lưu {len(data)} rows")
            
    def run(self):
        """Chạy demo"""
        print("🚀 Starting Tksheet Demo...")
        self.root.mainloop()


def main():
    """Main function"""
    print("=" * 60)
    print("🧪 TKSHEET DEMO - GOOGLE SHEETS STYLE")
    print("=" * 60)
    
    if not TKSHEET_AVAILABLE:
        print("❌ Tksheet chưa được cài đặt!")
        print("Chạy lệnh sau để cài đặt:")
        print("pip install tksheet")
        print()
        input("Nhấn Enter để đóng...")
        return
        
    print("✅ Tksheet có sẵn - Khởi động demo...")
    demo = TksheetDemo()
    demo.run()


if __name__ == "__main__":
    main()