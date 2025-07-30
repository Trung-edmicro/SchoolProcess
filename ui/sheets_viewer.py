"""
Google Sheets Viewer Component với Tksheet
Hiển thị dữ liệu từ Google Sheets với giao diện giống Google Sheets sử dụng Tksheet
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path

# Thêm project root vào Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

try:
    import tksheet
    TKSHEET_AVAILABLE = True
except ImportError:
    # Graceful fallback when tksheet is not installed
    tksheet = None  # type: ignore
    TKSHEET_AVAILABLE = False
    print("⚠️ Module tksheet chưa được cài đặt. Chạy: pip install tksheet")

from config.config_manager import get_config
from extractors.sheets_extractor import GoogleSheetsExtractor


class GoogleSheetsViewer:
    """Component hiển thị Google Sheets data với Tksheet"""
    
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame
        self.config = get_config()
        self.extractor = None
        self.data = []
        self.sheet_widget = None
        self.filtered_data = []
        self.setup_ui()
        
    def setup_ui(self):
        """Thiết lập giao diện với Tksheet"""
        # Configure parent frame
        self.parent_frame.configure(relief='flat', borderwidth=0)
        
        # Main container - luôn tạo trước
        self.main_frame = ttk.Frame(self.parent_frame)
        self.main_frame.pack(fill='both', expand=True, padx=0, pady=0)
        
        if not TKSHEET_AVAILABLE:
            self.show_tksheet_error()
            return
        
        # Header frame với Google colors
        self.create_header()
        
        # Separator line
        separator = ttk.Separator(self.main_frame, orient='horizontal')
        separator.pack(fill='x', pady=(0, 1))
        
        # Tksheet container
        self.create_tksheet_table()
        
        # Status frame
        self.create_status_bar()
        
        # Load data initially
        self.load_sheets_data()
        
    def show_tksheet_error(self):
        """Hiển thị lỗi khi tksheet chưa cài đặt"""
        error_frame = ttk.Frame(self.main_frame)
        error_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        error_label = ttk.Label(error_frame,
                               text="⚠️ Module 'tksheet' chưa được cài đặt",
                               font=('Segoe UI', 16, 'bold'),
                               foreground='red')
        error_label.pack(pady=20)
        
        instruction_label = ttk.Label(error_frame,
                                     text="Vui lòng chạy lệnh sau để cài đặt:\npip install tksheet",
                                     font=('Segoe UI', 12),
                                     justify='center')
        instruction_label.pack(pady=10)
        
        # Thêm button để install tksheet
        install_btn = ttk.Button(error_frame,
                                text="📦 Cài đặt Tksheet",
                                command=self.install_tksheet)
        install_btn.pack(pady=10)
        
    def install_tksheet(self):
        """Cài đặt tksheet module"""
        try:
            import subprocess
            import sys
            messagebox.showinfo("Cài đặt", "Đang cài đặt tksheet...\nVui lòng đợi...")
            result = subprocess.run([sys.executable, "-m", "pip", "install", "tksheet>=7.5.0"], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                messagebox.showinfo("Thành công", "Cài đặt tksheet thành công!\nVui lòng restart ứng dụng.")
            else:
                messagebox.showerror("Lỗi", f"Không thể cài đặt tksheet:\n{result.stderr}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi cài đặt: {str(e)}")
        
    def create_header(self):
        """Tạo header với info và controls giống Google Sheets"""
        # Header container với padding
        header_container = ttk.Frame(self.main_frame)
        header_container.pack(fill='x', padx=15, pady=(15, 10))
        
        # Title row
        title_row = ttk.Frame(header_container)
        title_row.pack(fill='x', pady=(0, 8))
        
        # Main title với Google-style icon
        title_frame = ttk.Frame(title_row)
        title_frame.pack(side='left')
        
        # Icon và title
        icon_label = ttk.Label(title_frame, 
                              text="📊", 
                              font=('Segoe UI', 16))
        icon_label.pack(side='left', padx=(0, 8))
        
        title_label = ttk.Label(title_frame,
                               text="School Process Data - Google Sheets View",
                               font=('Segoe UI', 16, 'bold'),
                               foreground='#1a73e8')
        title_label.pack(side='left')
        
        # Sheet ID info
        sheet_id = self.config.get_google_config().get('test_sheet_id', '')
        sheet_info_frame = ttk.Frame(title_row)
        sheet_info_frame.pack(side='right')
        
        ttk.Label(sheet_info_frame,
                 text="Sheet ID:",
                 font=('Segoe UI', 9),
                 foreground='#5f6368').pack(side='left', padx=(0, 5))
                 
        sheet_id_display = sheet_id[:20] + "..." if len(sheet_id) > 20 else sheet_id
        ttk.Label(sheet_info_frame,
                 text=sheet_id_display,
                 font=('Segoe UI', 9, 'bold'),
                 foreground='#137333').pack(side='left')
        
        # Toolbar row giống Google Sheets
        toolbar_row = ttk.Frame(header_container)
        toolbar_row.pack(fill='x')
        
        # Left toolbar
        left_toolbar = ttk.Frame(toolbar_row)
        left_toolbar.pack(side='left')
        
        # Action buttons với Google style
        self.refresh_btn = ttk.Button(left_toolbar, 
                                     text="🔄 Refresh",
                                     command=self.refresh_data)
        self.refresh_btn.pack(side='left', padx=(0, 8))
        
        self.add_btn = ttk.Button(left_toolbar,
                                 text="➕ Add Row", 
                                 command=self.add_row)
        self.add_btn.pack(side='left', padx=(0, 8))
        
        self.save_btn = ttk.Button(left_toolbar,
                                  text="💾 Save Changes",
                                  command=self.save_changes)
        self.save_btn.pack(side='left', padx=(0, 8))
        
        self.export_btn = ttk.Button(left_toolbar,
                                    text="📤 Export",
                                    command=self.export_data)
        self.export_btn.pack(side='left')
        
        # Right toolbar - Search
        right_toolbar = ttk.Frame(toolbar_row)
        right_toolbar.pack(side='right')
        
        search_frame = ttk.Frame(right_toolbar)
        search_frame.pack(side='right')
        
        # Search với Google style
        search_icon = ttk.Label(search_frame, text="🔍", font=('Segoe UI', 10))
        search_icon.pack(side='left', padx=(0, 5))
        
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.on_search)
        self.search_entry = ttk.Entry(search_frame, 
                                textvariable=self.search_var, 
                                width=25,
                                font=('Segoe UI', 10))
        self.search_entry.pack(side='left')
        self.search_entry.insert(0, "Search data...")
        
        # Bind events cho search entry
        self.search_entry.bind('<FocusIn>', self.on_search_focus_in)
        self.search_entry.bind('<FocusOut>', self.on_search_focus_out)
        
    def create_tksheet_table(self):
        """Tạo bảng Tksheet giống Google Sheets"""
        if not TKSHEET_AVAILABLE:
            return
            
        # Table container với margin giống Google Sheets
        table_container = ttk.Frame(self.main_frame)
        table_container.pack(fill='both', expand=True, padx=15, pady=(0, 10))
        
        # Tạo Tksheet widget
        self.sheet_widget = tksheet.Sheet(
            table_container,
            page_up_down_select_row=True,
            expand_sheet_if_paste_too_big=True,
            empty_horizontal=0,
            empty_vertical=0,
            show_horizontal_grid=True,
            show_vertical_grid=True,
            width=1000,
            height=500
        )
        
        # Configure Google Sheets styling
        self.configure_sheet_styling()
        
        # Pack the sheet
        self.sheet_widget.pack(fill='both', expand=True)
        
        # Bind events
        self.setup_sheet_bindings()
        
    def configure_sheet_styling(self):
        """Cấu hình styling giống Google Sheets"""
        if not self.sheet_widget:
            return
            
        # Google Sheets colors
        self.sheet_widget.set_options(
            font=('Segoe UI', 11, 'normal'),
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
        
        # Enable các tính năng editing
        self.sheet_widget.enable_bindings([
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
        
    def setup_sheet_bindings(self):
        """Thiết lập event bindings cho sheet"""
        if not self.sheet_widget:
            return
            
        # Bind cell edit events
        self.sheet_widget.extra_bindings([
            ("cell_select", self.on_cell_select),
            ("begin_edit_cell", self.on_begin_edit),
            ("end_edit_cell", self.on_end_edit),
            ("right_click_popup_menu", self.on_right_click)
        ])
        
    def setup_headers_and_columns(self):
        """Thiết lập headers và columns"""
        if not self.sheet_widget:
            return
            
        # Định nghĩa headers
        headers = [
            "STT",
            "ID Trường", 
            "Admin",
            "Mật khẩu",
            "Link Driver Dữ liệu",
            "SL GV nạp",
            "SL HS nạp", 
            "Notes"
        ]
        
        # Set headers
        self.sheet_widget.headers(headers)
        
        # Set column widths tối ưu
        column_widths = [60, 130, 150, 120, 250, 100, 100, 180]
        for i, width in enumerate(column_widths):
            self.sheet_widget.column_width(column=i, width=width)
        
    def create_status_bar(self):
        """Tạo status bar giống Google Sheets"""
        status_container = ttk.Frame(self.main_frame)
        status_container.pack(fill='x', padx=15, pady=(5, 15))
        
        # Separator trên status bar
        separator = ttk.Separator(status_container, orient='horizontal')
        separator.pack(fill='x', pady=(0, 8))
        
        status_frame = ttk.Frame(status_container)
        status_frame.pack(fill='x')
        
        # Status message
        self.status_var = tk.StringVar(value="📊 Sẵn sàng tải dữ liệu")
        status_label = ttk.Label(status_frame, 
                                textvariable=self.status_var,
                                font=('Segoe UI', 9),
                                foreground='#5f6368')
        status_label.pack(side='left')
        
        # Right side info
        right_info = ttk.Frame(status_frame)
        right_info.pack(side='right')
        
        # Row count
        self.row_count_var = tk.StringVar(value="0 rows")
        row_count_label = ttk.Label(right_info, 
                                   textvariable=self.row_count_var,
                                   font=('Segoe UI', 9, 'bold'),
                                   foreground='#137333')
        row_count_label.pack(side='right', padx=(10, 0))
        
        # Last updated
        self.last_updated_var = tk.StringVar(value="")
        updated_label = ttk.Label(right_info,
                                 textvariable=self.last_updated_var,
                                 font=('Segoe UI', 9),
                                 foreground='#5f6368')
        updated_label.pack(side='right', padx=(10, 0))
        
    def load_sheets_data(self):
        """Load dữ liệu từ Google Sheets"""
        if not TKSHEET_AVAILABLE:
            return
            
        try:
            self.status_var.set("🔄 Đang tải dữ liệu từ Google Sheets...")
            self.refresh_btn.config(state='disabled')
            
            # Get sheet config
            google_config = self.config.get_google_config()
            sheet_id = google_config.get('test_sheet_id')
            
            if not sheet_id:
                self.status_var.set("❌ Không tìm thấy GOOGLE_TEST_SHEET_ID trong config!")
                messagebox.showerror("Lỗi", "Không tìm thấy GOOGLE_TEST_SHEET_ID trong config!")
                return
                
            # Initialize extractor
            self.extractor = GoogleSheetsExtractor()
            
            # Định nghĩa columns cần extract
            required_columns = [
                'Tên trường',
                'Admin', 
                'Mật khẩu',
                'Link driver dữ liệu',
                'Số lượng GV nạp',
                'Số lượng HS nạp',
                'Notes'
            ]
            
            # Get data from sheets
            sheet_data = self.extractor.extract_required_columns(
                sheet_id=sheet_id,
                required_columns=required_columns
            )
            
            if not sheet_data:
                self.status_var.set("⚠️ Không có dữ liệu trong sheet")
                self.row_count_var.set("0 rows")
                return
                
            # Setup headers và load data
            self.setup_headers_and_columns()
            self.load_data_to_sheet(sheet_data)
            
            # Update status
            from datetime import datetime
            current_time = datetime.now().strftime("%H:%M:%S")
            self.status_var.set("✅ Tải dữ liệu thành công")
            self.last_updated_var.set(f"Updated: {current_time}")
            
        except Exception as e:
            error_msg = str(e)
            self.status_var.set(f"❌ Lỗi: {error_msg}")
            messagebox.showerror("Lỗi tải dữ liệu", 
                               f"Không thể tải dữ liệu từ Google Sheets:\n{error_msg}")
        finally:
            self.refresh_btn.config(state='normal')
            
    def load_data_to_sheet(self, sheet_data):
        """Load dữ liệu vào Tksheet"""
        if not self.sheet_widget or not sheet_data:
            return
            
        # Clear existing data
        self.sheet_widget.set_sheet_data([[]])
        
        # Parse data
        if 'data' not in sheet_data:
            self.row_count_var.set("0 rows")
            return
            
        rows = sheet_data['data']
        if not rows:
            self.row_count_var.set("0 rows")
            return
        
        # Convert data to table format
        table_data = []
        self.data = []
        
        for i, row in enumerate(rows):
            row_data = [
                i + 1,  # STT
                row.get('Tên trường', '') or '',
                row.get('Admin', '') or '',
                row.get('Mật khẩu', '') or '',
                row.get('Link driver dữ liệu', '') or '',
                row.get('Số lượng GV nạp', '') or '',
                row.get('Số lượng HS nạp', '') or '',
                row.get('Notes', '') or ''
            ]
            table_data.append(row_data)
            self.data.append(row_data)
        
        # Set data to sheet
        self.sheet_widget.set_sheet_data(table_data)
        self.filtered_data = table_data.copy()
        
        # Update row count
        self.row_count_var.set(f"{len(table_data)} rows")
        
        # Auto-fit columns
        self.auto_fit_columns()
        
    def auto_fit_columns(self):
        """Tự động điều chỉnh độ rộng cột"""
        if not self.sheet_widget:
            return
            
        try:
            # Auto-fit các cột quan trọng
            for col in range(self.sheet_widget.get_total_columns()):
                self.sheet_widget.set_column_width(column=col, width="displayed_text")
        except:
            pass  # Fallback nếu auto-fit không thành công
        
    def refresh_data(self):
        """Refresh dữ liệu"""
        self.load_sheets_data()
        
    def add_row(self):
        """Thêm row mới"""
        if not self.sheet_widget:
            messagebox.showwarning("Thông báo", "Sheet chưa được khởi tạo")
            return
            
        # Thêm row trống
        new_row = ["", "", "", "", "", "", "", ""]
        current_data = self.sheet_widget.get_sheet_data()
        current_data.append(new_row)
        self.sheet_widget.set_sheet_data(current_data)
        
        # Update row count
        self.row_count_var.set(f"{len(current_data)} rows")
        
        # Focus vào row mới
        new_row_index = len(current_data) - 1
        self.sheet_widget.select_row(new_row_index)
        self.sheet_widget.see(row=new_row_index, column=0)
        
    def save_changes(self):
        """Lưu thay đổi"""
        if not self.sheet_widget:
            return
            
        try:
            # Get current data from sheet
            current_data = self.sheet_widget.get_sheet_data()
            
            messagebox.showinfo("Lưu thay đổi", 
                              f"Đã lưu {len(current_data)} rows.\n\n"
                              "Lưu ý: Đây là demo - dữ liệu chưa được đồng bộ về Google Sheets.")
                              
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lưu thay đổi:\n{str(e)}")
            
    def export_data(self):
        """Export dữ liệu"""
        if not self.sheet_widget:
            return
            
        try:
            from tkinter import filedialog
            import csv
            
            # Chọn file để save
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Xuất dữ liệu"
            )
            
            if file_path:
                # Get data và headers
                data = self.sheet_widget.get_sheet_data()
                headers = ["STT", "ID Trường", "Admin", "Mật khẩu", 
                          "Link Driver Dữ liệu", "SL GV nạp", "SL HS nạp", "Notes"]
                
                # Write to CSV
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(headers)
                    writer.writerows(data)
                    
                messagebox.showinfo("Xuất dữ liệu", f"Đã xuất dữ liệu thành công!\nFile: {file_path}")
                
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất dữ liệu:\n{str(e)}")
        
    def on_search(self, *args):
        """Xử lý search với Tksheet"""
        if not self.sheet_widget:
            return
            
        search_term = self.search_var.get().lower()
        
        # Nếu search term là placeholder thì bỏ qua
        if search_term == "search data...":
            return
            
        if not search_term:
            # Hiển thị lại tất cả data
            self.sheet_widget.set_sheet_data(self.data)
            self.row_count_var.set(f"{len(self.data)} rows")
            return
            
        # Filter data
        filtered_data = []
        for row in self.data:
            # Check if search term matches any column
            match = any(search_term in str(cell).lower() for cell in row)
            if match:
                filtered_data.append(row)
                
        # Update sheet với filtered data
        self.sheet_widget.set_sheet_data(filtered_data)
        self.filtered_data = filtered_data
        
        # Update row count
        total_rows = len(self.data)
        filtered_rows = len(filtered_data)
        self.row_count_var.set(f"{filtered_rows} rows (filtered from {total_rows})")
        
    def on_search_focus_in(self, event):
        """Xử lý khi focus vào search entry"""
        if self.search_entry.get() == "Search data...":
            self.search_entry.delete(0, tk.END)
            
    def on_search_focus_out(self, event):
        """Xử lý khi focus ra khỏi search entry"""
        if not self.search_entry.get():
            self.search_entry.insert(0, "Search data...")
            
    def on_cell_select(self, event):
        """Xử lý khi select cell"""
        # Update status với thông tin cell hiện tại
        try:
            row = event.row
            col = event.column
            if row is not None and col is not None:
                headers = ["STT", "ID Trường", "Admin", "Mật khẩu", 
                          "Link Driver Dữ liệu", "SL GV nạp", "SL HS nạp", "Notes"]
                col_name = headers[col] if col < len(headers) else f"Column {col}"
                self.status_var.set(f"📍 Selected: {col_name} - Row {row + 1}")
        except:
            pass
            
    def on_begin_edit(self, event):
        """Xử lý khi bắt đầu edit cell"""
        try:
            row = event.row
            col = event.column
            if row is not None and col is not None:
                self.status_var.set(f"✏️ Editing cell Row {row + 1}, Column {col + 1}")
        except:
            pass
            
    def on_end_edit(self, event):
        """Xử lý khi kết thúc edit cell"""
        try:
            row = event.row
            col = event.column 
            if row is not None and col is not None:
                # Cập nhật data backup
                current_data = self.sheet_widget.get_sheet_data()
                if search_term := self.search_var.get().lower():
                    if search_term != "search data...":
                        # Nếu đang filter thì cập nhật cả original data
                        self.update_original_data_after_edit(row, col, current_data[row][col])
                else:
                    self.data = current_data.copy()
                    
                self.status_var.set(f"✅ Saved changes to Row {row + 1}, Column {col + 1}")
        except:
            pass
            
    def update_original_data_after_edit(self, filtered_row, col, new_value):
        """Cập nhật original data sau khi edit trong filtered view"""
        try:
            # Tìm row tương ứng trong original data
            filtered_data = self.sheet_widget.get_sheet_data()
            if filtered_row < len(filtered_data):
                row_data = filtered_data[filtered_row]
                # Tìm trong original data dựa trên STT hoặc unique identifier
                stt = row_data[0]  # STT column
                for i, orig_row in enumerate(self.data):
                    if orig_row[0] == stt:
                        self.data[i][col] = new_value
                        break
        except:
            pass
            
    def on_right_click(self, event):
        """Xử lý right click menu"""
        try:
            # Tạo context menu đơn giản
            menu = tk.Menu(self.sheet_widget, tearoff=0)
            menu.add_command(label="📋 Copy", command=lambda: self.sheet_widget.copy())
            menu.add_command(label="📄 Paste", command=lambda: self.sheet_widget.paste())
            menu.add_separator()
            menu.add_command(label="➕ Insert Row Above", command=self.insert_row_above)
            menu.add_command(label="➕ Insert Row Below", command=self.insert_row_below)
            menu.add_separator()
            menu.add_command(label="❌ Delete Row", command=self.delete_selected_row)
            
            # Show menu
            menu.tk_popup(event.x_root, event.y_root)
        except:
            pass
            
    def insert_row_above(self):
        """Chèn row trống phía trên row hiện tại"""
        try:
            selected_rows = self.sheet_widget.get_selected_rows()
            if selected_rows:
                row_index = min(selected_rows)
                new_row = ["", "", "", "", "", "", "", ""]
                self.sheet_widget.insert_row(row_index, new_row)
                self.update_row_count()
        except:
            pass
            
    def insert_row_below(self):
        """Chèn row trống phía dưới row hiện tại"""
        try:
            selected_rows = self.sheet_widget.get_selected_rows()
            if selected_rows:
                row_index = max(selected_rows) + 1
                new_row = ["", "", "", "", "", "", "", ""]
                self.sheet_widget.insert_row(row_index, new_row)
                self.update_row_count()
        except:
            pass
            
    def delete_selected_row(self):
        """Xóa row được chọn"""
        try:
            selected_rows = self.sheet_widget.get_selected_rows()
            if selected_rows:
                # Confirm deletion
                if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa {len(selected_rows)} row(s)?"):
                    # Delete from highest index to lowest to avoid index shifting
                    for row_index in sorted(selected_rows, reverse=True):
                        self.sheet_widget.delete_row(row_index)
                    self.update_row_count()
        except:
            pass
            
    def update_row_count(self):
        """Cập nhật row count sau khi thêm/xóa row"""
        try:
            current_data = self.sheet_widget.get_sheet_data()
            self.row_count_var.set(f"{len(current_data)} rows")
        except:
            pass
            
    def get_selected_row_data(self):
        """Lấy dữ liệu từ row được chọn để xử lý workflow"""
        try:
            if not self.sheet_widget:
                return None
                
            selected_rows = self.sheet_widget.get_selected_rows()
            if not selected_rows:
                return None
                
            # Lấy row đầu tiên được chọn
            row_index = min(selected_rows)
            current_data = self.sheet_widget.get_sheet_data()
            
            if row_index >= len(current_data):
                return None
                
            row_data = current_data[row_index]
            
            # Convert sang dictionary format như extractor trả về
            school_data = {
                'STT': row_data[0] if len(row_data) > 0 else '',
                'Tên trường': row_data[1] if len(row_data) > 1 else '',
                'Admin': row_data[2] if len(row_data) > 2 else '',
                'Mật khẩu': row_data[3] if len(row_data) > 3 else '',
                'Link driver dữ liệu': row_data[4] if len(row_data) > 4 else '',
                'Số lượng GV nạp': row_data[5] if len(row_data) > 5 else '',
                'Số lượng HS nạp': row_data[6] if len(row_data) > 6 else '',
                'Notes': row_data[7] if len(row_data) > 7 else ''
            }
            
            return school_data
            
        except Exception as e:
            print(f"Error getting selected row data: {e}")
            return None
            
    def get_selected_row_info(self):
        """Lấy thông tin về row được chọn (cho hiển thị)"""
        try:
            if not self.sheet_widget:
                return "Không có sheet data"
                
            selected_rows = self.sheet_widget.get_selected_rows()
            if not selected_rows:
                return "Chưa chọn row nào"
                
            row_index = min(selected_rows)
            school_data = self.get_selected_row_data()
            
            if school_data:
                school_name = school_data.get('Tên trường', 'N/A')
                admin_email = school_data.get('Admin', 'N/A')
                return f"Row {row_index + 1}: {school_name} ({admin_email})"
            else:
                return f"Row {row_index + 1}: Dữ liệu không hợp lệ"
                
        except Exception as e:
            return f"Lỗi: {str(e)}"