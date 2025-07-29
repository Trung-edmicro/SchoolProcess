"""
Google Sheets Viewer Component
Hiển thị dữ liệu từ Google Sheets với giao diện giống Google Sheets
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path

# Thêm project root vào Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from config.config_manager import get_config
from extractors.sheets_extractor import GoogleSheetsExtractor


class GoogleSheetsViewer:
    """Component hiển thị Google Sheets data"""
    
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame
        self.config = get_config()
        self.extractor = None
        self.data = []
        self.setup_ui()
        
    def setup_ui(self):
        """Thiết lập giao diện"""
        # Configure parent frame để giống Google Sheets
        self.parent_frame.configure(relief='flat', borderwidth=0)
        
        # Main container với background trắng
        self.main_frame = ttk.Frame(self.parent_frame)
        self.main_frame.pack(fill='both', expand=True, padx=0, pady=0)
        
        # Header frame với Google colors
        self.create_header()
        
        # Separator line giống Google Sheets
        separator = ttk.Separator(self.main_frame, orient='horizontal')
        separator.pack(fill='x', pady=(0, 1))
        
        # Sheets table frame
        self.create_sheets_table()
        
        # Status frame
        self.create_status_bar()
        
        # Load data initially
        self.load_sheets_data()
        
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
                               text="School Process Data",
                               font=('Segoe UI', 16, 'bold'),
                               foreground='#1a73e8')
        title_label.pack(side='left')
        
        # Sheet ID info với Google style
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
                                 text="➕ Add", 
                                 command=self.add_row)
        self.add_btn.pack(side='left', padx=(0, 8))
        
        self.save_btn = ttk.Button(left_toolbar,
                                  text="💾 Save",
                                  command=self.save_changes,
                                  state='disabled')
        self.save_btn.pack(side='left')
        
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
        search_entry = ttk.Entry(search_frame, 
                                textvariable=self.search_var, 
                                width=25,
                                font=('Segoe UI', 10))
        search_entry.pack(side='left')
        search_entry.insert(0, "Search data...")
        
    def create_sheets_table(self):
        """Tạo bảng hiển thị dữ liệu giống Google Sheets"""
        # Table container với margin giống Google Sheets
        table_container = ttk.Frame(self.main_frame)
        table_container.pack(fill='both', expand=True, padx=15, pady=(0, 10))
        
        # Định nghĩa columns với widths tối ưu
        self.columns = [
            ("stt", "STT", 60),
            ("id_truong", "ID Trường", 130),
            ("admin", "Admin", 150), 
            ("mat_khau", "Mật khẩu", 120),
            ("link_driver", "Link Driver Dữ liệu", 250),
            ("so_luong_gv", "SL GV nạp", 100),
            ("so_luong_hs", "SL HS nạp", 100),
            ("notes", "Notes", 180)
        ]
        
        # Treeview với Google Sheets styling
        style = ttk.Style()
        
        # Configure Treeview style giống Google Sheets
        style.configure("Sheets.Treeview",
                       background="white",
                       foreground="black",
                       rowheight=32,
                       fieldbackground="white",
                       borderwidth=1,
                       relief="solid")
        
        style.configure("Sheets.Treeview.Heading",
                       background="#f8f9fa",
                       foreground="#202124",
                       font=('Segoe UI', 10, 'bold'),
                       borderwidth=1,
                       relief="solid")
        
        self.tree = ttk.Treeview(table_container, 
                                columns=[col[0] for col in self.columns],
                                show='headings',
                                height=18,
                                style="Sheets.Treeview")
        
        # Configure columns với Google Sheets headers
        for col_id, col_text, col_width in self.columns:
            self.tree.heading(col_id, text=col_text, 
                             command=lambda c=col_id: self.sort_column(c))
            self.tree.column(col_id, width=col_width, minwidth=80, anchor='w')
            
        # Scrollbars với Google style
        v_scrollbar = ttk.Scrollbar(table_container, orient='vertical', command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(table_container, orient='horizontal', command=self.tree.xview)
        
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)
        
        # Bind events
        self.tree.bind('<Double-1>', self.on_double_click)
        self.tree.bind('<Button-3>', self.on_right_click)
        self.tree.bind('<Button-1>', self.on_select)
        
        # Configure row colors giống Google Sheets
        self.tree.tag_configure('oddrow', background='#ffffff')
        self.tree.tag_configure('evenrow', background='#f8f9fa')
        self.tree.tag_configure('selected', background='#e8f0fe', foreground='#1a73e8')
        self.tree.tag_configure('hover', background='#f1f3f4')
        
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
        
        # Row count với Google style
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
                'Tên trường',  # ID Trường
                'Admin',       # Admin
                'Mật khẩu',    # Mật khẩu
                'Link driver dữ liệu',  # Link driver
                'Số lượng GV nạp',  # Số lượng GV (optional)
                'Số lượng HS nạp',  # Số lượng HS (optional)
                'Notes'        # Notes (optional)
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
                
            # Parse và hiển thị data
            self.parse_and_display_data(sheet_data)
            
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
            
    def parse_and_display_data(self, sheet_data):
        """Parse và hiển thị dữ liệu lên bảng"""
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        self.data = []
        
        # Kiểm tra cấu trúc dữ liệu trả về từ extractor
        if not sheet_data or 'data' not in sheet_data:
            self.row_count_var.set("0 rows")
            self.status_var.set("❌ Không có dữ liệu hoặc cấu trúc dữ liệu không đúng")
            return
            
        rows = sheet_data['data']
        metadata = sheet_data.get('metadata', {})
        
        if not rows:
            self.row_count_var.set("0 rows")
            self.status_var.set("⚠️ Sheet không có dữ liệu")
            return
        
        # Mapping từ columns extractor về columns hiển thị
        extractor_to_display = {
            'Tên trường': 'id_truong',
            'Admin': 'admin',
            'Mật khẩu': 'mat_khau',
            'Link driver dữ liệu': 'link_driver'
        }
        
        # Process data từ extractor
        for i, row in enumerate(rows):
            # Extract data theo mapping
            row_data = {
                'stt': i + 1,
                'id_truong': row.get('Tên trường', '') or '',
                'admin': row.get('Admin', '') or '',
                'mat_khau': row.get('Mật khẩu', '') or '',  # Sẽ được map từ 'Password'
                'link_driver': row.get('Link driver dữ liệu', '') or '',
                'so_luong_gv': row.get('Số lượng GV nạp', '') or '',  
                'so_luong_hs': row.get('Số lượng HS nạp', '') or '',  
                'notes': row.get('Notes', '') or ''  
            }
            
            self.data.append(row_data)
            
            # Insert vào tree với alternating colors
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            self.tree.insert('', 'end', values=list(row_data.values()), tags=(tag,))
            
        # Update metadata info
        found_columns = metadata.get('found_columns', {})
        missing_columns = metadata.get('missing_columns', [])
        
        status_msg = f"✅ Tải {len(rows)} rows"
        if found_columns:
            status_msg += f" | Found: {', '.join(found_columns.keys())}"
        if missing_columns:
            status_msg += f" | Missing: {', '.join(missing_columns)}"
            
        self.status_var.set(status_msg)
        self.row_count_var.set(f"{len(rows)} rows")
        
    def create_header_mapping(self, headers):
        """Tạo mapping từ headers thực tế sang columns của chúng ta"""
        # Mapping keywords to find relevant columns
        mapping_keywords = {
            'id_truong': ['id', 'truong', 'school', 'mã'],
            'admin': ['admin', 'user', 'tài khoản'],
            'mat_khau': ['password', 'mật khẩu', 'pass'],
            'link_driver': ['link', 'drive', 'driver', 'url'],
            'so_luong_gv': ['gv', 'giáo viên', 'teacher', 'số lượng gv'],
            'so_luong_hs': ['hs', 'học sinh', 'student', 'số lượng hs'],
            'notes': ['note', 'ghi chú', 'comment', 'remark']
        }
        
        header_mapping = {}
        
        for col_key, keywords in mapping_keywords.items():
            for i, header in enumerate(headers):
                header_lower = str(header).lower()
                if any(keyword in header_lower for keyword in keywords):
                    header_mapping[col_key] = i
                    break
                    
        return header_mapping
        
    def extract_row_data(self, row, header_mapping, stt):
        """Extract data từ row theo mapping"""
        row_data = {
            'stt': stt,
            'id_truong': '',
            'admin': '',
            'mat_khau': '',
            'link_driver': '',
            'so_luong_gv': '',
            'so_luong_hs': '',
            'notes': ''
        }
        
        # Fill data theo mapping
        for col_key, col_index in header_mapping.items():
            if col_index < len(row):
                row_data[col_key] = str(row[col_index]) if row[col_index] else ''
                
        return row_data
        
    def refresh_data(self):
        """Refresh dữ liệu"""
        self.load_sheets_data()
        
    def add_row(self):
        """Thêm row mới"""
        # TODO: Implement add row functionality
        messagebox.showinfo("Chức năng", "Chức năng thêm row sẽ được triển khai sau")
        
    def save_changes(self):
        """Lưu thay đổi"""
        # TODO: Implement save functionality  
        messagebox.showinfo("Chức năng", "Chức năng lưu thay đổi sẽ được triển khai sau")
        
    def on_search(self, *args):
        """Xử lý search"""
        search_term = self.search_var.get().lower()
        
        # Clear current display
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Filter và hiển thị data
        filtered_count = 0
        for i, row_data in enumerate(self.data):
            # Check if search term matches any column
            match = any(search_term in str(value).lower() for value in row_data.values())
            
            if not search_term or match:
                tag = 'evenrow' if filtered_count % 2 == 0 else 'oddrow'
                self.tree.insert('', 'end', values=list(row_data.values()), tags=(tag,))
                filtered_count += 1
                
        self.row_count_var.set(f"{filtered_count} rows" + (f" (filtered from {len(self.data)})" if search_term else ""))
        
    def on_double_click(self, event):
        """Xử lý double click để edit"""
        # TODO: Implement edit functionality
        item = self.tree.selection()[0] if self.tree.selection() else None
        if item:
            values = self.tree.item(item, 'values')
            messagebox.showinfo("Edit Row", f"Chức năng edit sẽ được triển khai sau\nRow data: {values[0]} - {values[1]}")
            
    def on_select(self, event):
        """Xử lý selection event"""
        selected_items = self.tree.selection()
        if selected_items:
            # Highlight selected row
            for item in selected_items:
                self.tree.set(item, values=self.tree.item(item, 'values'))
            
    def on_right_click(self, event):
        """Xử lý right click menu"""
        # TODO: Implement context menu với các options như Google Sheets
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            # Context menu sẽ được implement sau
            
    def sort_column(self, col):
        """Sort column"""
        # TODO: Implement sorting như Google Sheets
        messagebox.showinfo("Sort", f"Chức năng sort column '{col}' sẽ được triển khai sau")
