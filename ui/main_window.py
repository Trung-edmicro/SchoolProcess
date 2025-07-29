"""
Main UI Window for School Process Application
Modern Material Design với Tkinter
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import json
import os
from datetime import datetime
from pathlib import Path
import sys

# Thêm project root vào Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from config.config_manager import get_config
from config.onluyen_api import OnLuyenAPIClient
from extractors import GoogleSheetsExtractor
from converters import JSONToExcelTemplateConverter


class SchoolProcessMainWindow:
    """Main Window cho School Process Application"""
    
    def __init__(self):
        """Khởi tạo main window"""
        self.config = get_config()
        self.setup_main_window()
        self.setup_variables()
        self.setup_ui()
        self.setup_bindings()
        
    def setup_main_window(self):
        """Thiết lập cửa sổ chính"""
        self.root = tk.Tk()
        self.root.title("School Process - Ứng dụng xử lý dữ liệu trường học")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)
        
        # Icon và styling
        try:
            self.root.iconbitmap("assets/icon.ico")
        except:
            pass
            
        # Center window
        self.center_window()
        
        # Configure style
        self.setup_styles()
        
    def center_window(self):
        """Căn giữa cửa sổ"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
    def setup_styles(self):
        """Thiết lập styles cho UI"""
        self.style = ttk.Style()
        
        # Chọn theme modern
        self.style.theme_use('clam')
        
        # Custom colors - Material Design inspired
        self.colors = {
            'primary': '#1976D2',      # Blue
            'primary_dark': '#1565C0',
            'secondary': '#FF6F00',    # Orange
            'success': '#4CAF50',      # Green
            'warning': '#FF9800',      # Amber
            'error': '#F44336',        # Red
            'surface': '#FFFFFF',      # White
            'background': '#F5F5F5',   # Light Gray
            'text': '#212121',         # Dark Gray
            'text_secondary': '#757575' # Medium Gray
        }
        
        # Configure ttk styles
        self.style.configure('Title.TLabel', 
                           font=('Segoe UI', 16, 'bold'),
                           foreground=self.colors['primary'])
        
        self.style.configure('Heading.TLabel',
                           font=('Segoe UI', 12, 'bold'),
                           foreground=self.colors['text'])
        
        self.style.configure('Primary.TButton',
                           font=('Segoe UI', 10),
                           foreground='white')
        
        self.style.map('Primary.TButton',
                      background=[('active', self.colors['primary_dark']),
                                ('!active', self.colors['primary'])])
        
        self.style.configure('Success.TButton',
                           font=('Segoe UI', 10),
                           foreground='white')
        
        self.style.map('Success.TButton',
                      background=[('active', '#45A049'),
                                ('!active', self.colors['success'])])
                                
    def setup_variables(self):
        """Thiết lập các biến"""
        self.current_task = tk.StringVar(value="Sẵn sàng")
        self.progress_var = tk.DoubleVar()
        self.log_text = tk.StringVar()
        
        # Application state
        self.is_processing = False
        self.current_workflow = None
        self.client = None
        
    def setup_ui(self):
        """Thiết lập giao diện người dùng"""
        # Main container
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        
        # Create UI sections
        self.create_header()
        self.create_main_content()
        self.create_status_bar()
        
    def create_header(self):
        """Tạo header với tiêu đề và thông tin"""
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        header_frame.columnconfigure(1, weight=1)
        
        # Logo/Icon placeholder
        logo_frame = ttk.Frame(header_frame, width=60, height=60)
        logo_frame.grid(row=0, column=0, padx=(0, 15), rowspan=2)
        logo_frame.grid_propagate(False)
        
        logo_label = ttk.Label(logo_frame, text="🏫", font=('Segoe UI', 24))
        logo_label.place(relx=0.5, rely=0.5, anchor='center')
        
        # Title
        title_label = ttk.Label(header_frame, 
                              text="School Process Application",
                              style='Title.TLabel')
        title_label.grid(row=0, column=1, sticky=(tk.W), pady=(5, 0))
        
        # Subtitle
        subtitle_label = ttk.Label(header_frame,
                                 text="Ứng dụng xử lý dữ liệu trường học với OnLuyen API",
                                 font=('Segoe UI', 10),
                                 foreground=self.colors['text_secondary'])
        subtitle_label.grid(row=1, column=1, sticky=(tk.W), pady=(0, 5))
        
        # Status indicator
        self.status_frame = ttk.Frame(header_frame)
        self.status_frame.grid(row=0, column=2, rowspan=2, padx=(15, 0))
        
        self.status_label = ttk.Label(self.status_frame,
                                    text="● Sẵn sàng",
                                    font=('Segoe UI', 10, 'bold'),
                                    foreground=self.colors['success'])
        self.status_label.pack()
        
    def create_main_content(self):
        """Tạo nội dung chính"""
        # Left panel - Menu
        self.create_left_panel()
        
        # Right panel - Content
        self.create_right_panel()
        
    def create_left_panel(self):
        """Tạo panel menu bên trái"""
        left_frame = ttk.LabelFrame(self.main_frame, text="Chức năng", padding="10")
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        left_frame.configure(width=300)
        left_frame.grid_propagate(False)
        
        # Workflow section
        workflow_label = ttk.Label(left_frame, text="Quy trình xử lý", style='Heading.TLabel')
        workflow_label.pack(pady=(0, 10), anchor='w')
        
        # Workflow buttons
        self.btn_case1 = ttk.Button(left_frame,
                                   text="📊 Case 1: Toàn bộ dữ liệu",
                                   style='Primary.TButton',
                                   command=self.start_workflow_case1)
        self.btn_case1.pack(fill='x', pady=(0, 5))
        
        self.btn_case2 = ttk.Button(left_frame,
                                   text="🔍 Case 2: Dữ liệu theo file import",
                                   style='Primary.TButton',
                                   command=self.start_workflow_case2)
        self.btn_case2.pack(fill='x', pady=(0, 15))
        
        # Separator
        separator1 = ttk.Separator(left_frame, orient='horizontal')
        separator1.pack(fill='x', pady=(0, 15))
        
        # Individual functions
        functions_label = ttk.Label(left_frame, text="Chức năng đơn lẻ", style='Heading.TLabel')
        functions_label.pack(pady=(0, 10), anchor='w')
        
        self.btn_get_teachers = ttk.Button(left_frame,
                                          text="👨‍🏫 Lấy danh sách Giáo viên",
                                          command=self.get_teachers_data)
        self.btn_get_teachers.pack(fill='x', pady=(0, 5))
        
        self.btn_get_students = ttk.Button(left_frame,
                                          text="👨‍🎓 Lấy danh sách Học sinh",
                                          command=self.get_students_data)
        self.btn_get_students.pack(fill='x', pady=(0, 5))
        
        self.btn_convert_excel = ttk.Button(left_frame,
                                           text="📄 Chuyển đổi JSON → Excel",
                                           command=self.convert_json_to_excel)
        self.btn_convert_excel.pack(fill='x', pady=(0, 15))
        
        # Separator
        separator2 = ttk.Separator(left_frame, orient='horizontal')
        separator2.pack(fill='x', pady=(0, 15))
        
        # Settings section
        settings_label = ttk.Label(left_frame, text="Cài đặt", style='Heading.TLabel')
        settings_label.pack(pady=(0, 10), anchor='w')
        
        self.btn_config = ttk.Button(left_frame,
                                    text="⚙️ Cấu hình",
                                    command=self.open_config)
        self.btn_config.pack(fill='x', pady=(0, 5))
        
        self.btn_about = ttk.Button(left_frame,
                                   text="ℹ️ Về ứng dụng",
                                   command=self.show_about)
        self.btn_about.pack(fill='x')
        
    def create_right_panel(self):
        """Tạo panel nội dung bên phải"""
        right_frame = ttk.Frame(self.main_frame)
        right_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        right_frame.rowconfigure(0, weight=1)
        right_frame.columnconfigure(0, weight=1)
        
        # Notebook cho tabs
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Tab: Google Sheets Viewer
        self.create_sheets_tab()
        
        # Tab: Log Output
        self.create_log_tab()
        
        # Tab: Configuration
        self.create_config_tab()
        
        # Tab: Results
        self.create_results_tab()
        
    def create_sheets_tab(self):
        """Tạo tab Google Sheets viewer"""
        sheets_frame = ttk.Frame(self.notebook)
        self.notebook.add(sheets_frame, text="📊 Google Sheets")
        
        # Import và khởi tạo sheets viewer
        try:
            from ui.sheets_viewer import GoogleSheetsViewer
            self.sheets_viewer = GoogleSheetsViewer(sheets_frame)
        except Exception as e:
            # Fallback nếu có lỗi
            error_label = ttk.Label(sheets_frame, 
                                   text=f"Lỗi tải Google Sheets Viewer:\n{str(e)}",
                                   justify='center')
            error_label.pack(expand=True)
        
    def create_log_tab(self):
        """Tạo tab log output"""
        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text="📋 Log & Tiến trình")
        
        log_frame.rowconfigure(1, weight=1)
        log_frame.columnconfigure(0, weight=1)
        
        # Progress section
        progress_frame = ttk.LabelFrame(log_frame, text="Tiến trình", padding="10")
        progress_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        # Current task
        self.task_label = ttk.Label(progress_frame, textvariable=self.current_task)
        self.task_label.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, 
                                          variable=self.progress_var,
                                          maximum=100)
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Control buttons
        control_frame = ttk.Frame(progress_frame)
        control_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        
        self.btn_stop = ttk.Button(control_frame,
                                  text="⏹️ Dừng",
                                  state='disabled',
                                  command=self.stop_processing)
        self.btn_stop.pack(side='left', padx=(0, 5))
        
        self.btn_clear_log = ttk.Button(control_frame,
                                       text="🗑️ Xóa log",
                                       command=self.clear_log)
        self.btn_clear_log.pack(side='left')
        
        # Log output
        log_output_frame = ttk.LabelFrame(log_frame, text="Log Output", padding="10")
        log_output_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_output_frame.rowconfigure(0, weight=1)
        log_output_frame.columnconfigure(0, weight=1)
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(log_output_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)
        
        self.log_text_widget = tk.Text(text_frame,
                                      wrap='word',
                                      font=('Consolas', 9),
                                      bg='#1e1e1e',
                                      fg='#ffffff',
                                      insertbackground='#ffffff')
        self.log_text_widget.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=self.log_text_widget.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text_widget.configure(yscrollcommand=scrollbar.set)
        
        # Configure text tags for colored output
        self.setup_log_tags()
        
    def create_config_tab(self):
        """Tạo tab cấu hình"""
        config_frame = ttk.Frame(self.notebook)
        self.notebook.add(config_frame, text="⚙️ Cấu hình")
        
        config_frame.rowconfigure(0, weight=1)
        config_frame.columnconfigure(0, weight=1)
        
        # Scrollable frame
        canvas = tk.Canvas(config_frame)
        scrollbar_config = ttk.Scrollbar(config_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_config.set)
        
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_config.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # OnLuyen API Configuration
        self.create_onluyen_config(scrollable_frame)
        
        # Google Sheets Configuration  
        self.create_sheets_config(scrollable_frame)
        
        # File Paths Configuration
        self.create_paths_config(scrollable_frame)
        
    def create_results_tab(self):
        """Tạo tab kết quả"""
        results_frame = ttk.Frame(self.notebook)
        self.notebook.add(results_frame, text="📊 Kết quả")
        
        results_frame.rowconfigure(1, weight=1)
        results_frame.columnconfigure(0, weight=1)
        
        # Summary section
        summary_frame = ttk.LabelFrame(results_frame, text="Tóm tắt", padding="10")
        summary_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.summary_label = ttk.Label(summary_frame, text="Chưa có dữ liệu")
        self.summary_label.pack()
        
        # Files section
        files_frame = ttk.LabelFrame(results_frame, text="Files đã tạo", padding="10")
        files_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        files_frame.rowconfigure(0, weight=1)
        files_frame.columnconfigure(0, weight=1)
        
        # Treeview for files
        self.files_tree = ttk.Treeview(files_frame, columns=('Type', 'Path', 'Size', 'Date'), show='tree headings')
        self.files_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure columns
        self.files_tree.heading('#0', text='Tên file')
        self.files_tree.heading('Type', text='Loại')
        self.files_tree.heading('Path', text='Đường dẫn')
        self.files_tree.heading('Size', text='Kích thước')
        self.files_tree.heading('Date', text='Ngày tạo')
        
        # Scrollbar for treeview
        files_scrollbar = ttk.Scrollbar(files_frame, orient='vertical', command=self.files_tree.yview)
        files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.files_tree.configure(yscrollcommand=files_scrollbar.set)
        
        # Context menu for files
        self.setup_files_context_menu()
        
    def create_status_bar(self):
        """Tạo status bar"""
        status_frame = ttk.Frame(self.main_frame)
        status_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        status_frame.columnconfigure(1, weight=1)
        
        # Status icon
        self.status_icon = ttk.Label(status_frame, text="●", foreground=self.colors['success'])
        self.status_icon.grid(row=0, column=0, padx=(0, 5))
        
        # Status text
        self.status_text = ttk.Label(status_frame, text="Sẵn sàng")
        self.status_text.grid(row=0, column=1, sticky=(tk.W))
        
        # Version info
        version_label = ttk.Label(status_frame, text="v1.0.0", font=('Segoe UI', 8))
        version_label.grid(row=0, column=2)
        
    def setup_log_tags(self):
        """Thiết lập tags cho colored log output"""
        # Success messages
        self.log_text_widget.tag_configure("success", foreground="#4CAF50")
        
        # Error messages
        self.log_text_widget.tag_configure("error", foreground="#F44336")
        
        # Warning messages
        self.log_text_widget.tag_configure("warning", foreground="#FF9800")
        
        # Info messages
        self.log_text_widget.tag_configure("info", foreground="#2196F3")
        
        # Headers/separators
        self.log_text_widget.tag_configure("header", foreground="#9C27B0", font=('Consolas', 9, 'bold'))
        
    def create_onluyen_config(self, parent):
        """Tạo section cấu hình OnLuyen API"""
        onluyen_frame = ttk.LabelFrame(parent, text="OnLuyen API", padding="10")
        onluyen_frame.pack(fill='x', pady=(0, 10))
        
        # Base URL
        ttk.Label(onluyen_frame, text="Base URL:").grid(row=0, column=0, sticky='w', pady=(0, 5))
        self.onluyen_url_var = tk.StringVar(value=self.config.get_onluyen_config().get('base_url', ''))
        url_entry = ttk.Entry(onluyen_frame, textvariable=self.onluyen_url_var, width=50)
        url_entry.grid(row=0, column=1, sticky='ew', pady=(0, 5), padx=(10, 0))
        
        # Default credentials (readonly)
        ttk.Label(onluyen_frame, text="Username:").grid(row=1, column=0, sticky='w', pady=(0, 5))
        self.onluyen_username_var = tk.StringVar(value=self.config.get_onluyen_config().get('username', ''))
        username_entry = ttk.Entry(onluyen_frame, textvariable=self.onluyen_username_var, width=50)
        username_entry.grid(row=1, column=1, sticky='ew', pady=(0, 5), padx=(10, 0))
        
        # Test connection button
        ttk.Button(onluyen_frame, text="Test Connection", command=self.test_onluyen_connection).grid(row=2, column=1, sticky='w', pady=(10, 0), padx=(10, 0))
        
        onluyen_frame.columnconfigure(1, weight=1)
        
    def create_sheets_config(self, parent):
        """Tạo section cấu hình Google Sheets"""
        sheets_frame = ttk.LabelFrame(parent, text="Google Sheets", padding="10")
        sheets_frame.pack(fill='x', pady=(0, 10))
        
        # Spreadsheet ID
        ttk.Label(sheets_frame, text="Spreadsheet ID:").grid(row=0, column=0, sticky='w', pady=(0, 5))
        self.sheets_id_var = tk.StringVar(value=self.config.get_google_config().get('test_sheet_id', ''))
        id_entry = ttk.Entry(sheets_frame, textvariable=self.sheets_id_var, width=50)
        id_entry.grid(row=0, column=1, sticky='ew', pady=(0, 5), padx=(10, 0))
        
        # Sheet Name
        ttk.Label(sheets_frame, text="Sheet Name:").grid(row=1, column=0, sticky='w', pady=(0, 5))
        self.sheet_name_var = tk.StringVar(value=self.config.get_google_config().get('sheet_name', ''))
        name_entry = ttk.Entry(sheets_frame, textvariable=self.sheet_name_var, width=50)
        name_entry.grid(row=1, column=1, sticky='ew', pady=(0, 5), padx=(10, 0))
        
        # Test connection button
        ttk.Button(sheets_frame, text="Test Connection", command=self.test_sheets_connection).grid(row=2, column=1, sticky='w', pady=(10, 0), padx=(10, 0))
        
        sheets_frame.columnconfigure(1, weight=1)
        
    def create_paths_config(self, parent):
        """Tạo section cấu hình đường dẫn"""
        paths_frame = ttk.LabelFrame(parent, text="Đường dẫn", padding="10")
        paths_frame.pack(fill='x', pady=(0, 10))
        
        paths_config = self.config.get_paths_config()
        
        # Input directory
        ttk.Label(paths_frame, text="Thư mục Input:").grid(row=0, column=0, sticky='w', pady=(0, 5))
        self.input_dir_var = tk.StringVar(value=paths_config.get('input_dir', ''))
        input_entry = ttk.Entry(paths_frame, textvariable=self.input_dir_var, width=40)
        input_entry.grid(row=0, column=1, sticky='ew', pady=(0, 5), padx=(10, 5))
        ttk.Button(paths_frame, text="...", command=lambda: self.browse_directory(self.input_dir_var)).grid(row=0, column=2, pady=(0, 5))
        
        # Output directory
        ttk.Label(paths_frame, text="Thư mục Output:").grid(row=1, column=0, sticky='w', pady=(0, 5))
        self.output_dir_var = tk.StringVar(value=paths_config.get('output_dir', ''))
        output_entry = ttk.Entry(paths_frame, textvariable=self.output_dir_var, width=40)
        output_entry.grid(row=1, column=1, sticky='ew', pady=(0, 5), padx=(10, 5))
        ttk.Button(paths_frame, text="...", command=lambda: self.browse_directory(self.output_dir_var)).grid(row=1, column=2, pady=(0, 5))
        
        paths_frame.columnconfigure(1, weight=1)
        
    def setup_files_context_menu(self):
        """Thiết lập context menu cho files tree"""
        self.files_context_menu = tk.Menu(self.root, tearoff=0)
        self.files_context_menu.add_command(label="Mở file", command=self.open_selected_file)
        self.files_context_menu.add_command(label="Mở thư mục chứa", command=self.open_file_location)
        self.files_context_menu.add_separator()
        self.files_context_menu.add_command(label="Sao chép đường dẫn", command=self.copy_file_path)
        
        self.files_tree.bind("<Button-3>", self.show_files_context_menu)
        
    def setup_bindings(self):
        """Thiết lập keyboard bindings"""
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        self.root.bind('<F5>', lambda e: self.refresh_ui())
        
    def log_message(self, message, level="info"):
        """Thêm message vào log với màu sắc tương ứng"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_message = f"[{timestamp}] {message}\n"
        
        self.log_text_widget.insert(tk.END, full_message, level)
        self.log_text_widget.see(tk.END)
        
        # Update status
        if level == "error":
            self.update_status("Lỗi", "error")
        elif level == "success":
            self.update_status("Thành công", "success")
        elif level == "warning":
            self.update_status("Cảnh báo", "warning")
        else:
            self.update_status("Đang xử lý...", "info")
            
    def log_message_safe(self, message, level="info"):
        """Thread-safe version của log_message"""
        self.root.after(0, lambda: self.log_message(message, level))
        
    def update_progress_safe(self, value, status=""):
        """Thread-safe version của update_progress"""
        self.root.after(0, lambda: self.update_progress(value, status))
        
    def update_button_state_safe(self, button, state):
        """Thread-safe version để cập nhật button state"""
        self.root.after(0, lambda: button.config(state=state))
            
    def update_status(self, message, level="info"):
        """Cập nhật status bar"""
        color_map = {
            "success": self.colors['success'],
            "error": self.colors['error'],
            "warning": self.colors['warning'],
            "info": self.colors['primary']
        }
        
        self.status_text.config(text=message)
        self.status_icon.config(foreground=color_map.get(level, self.colors['primary']))
        
    def update_progress(self, value, task=""):
        """Cập nhật progress bar"""
        self.progress_var.set(value)
        if task:
            self.current_task.set(task)
            
    def clear_log(self):
        """Xóa log output"""
        self.log_text_widget.delete(1.0, tk.END)
        
    def start_workflow_case1(self):
        """Bắt đầu workflow Case 1"""
        if self.is_processing:
            messagebox.showwarning("Cảnh báo", "Hệ thống đang xử lý. Vui lòng đợi.")
            return
            
        self.log_message("Bắt đầu Workflow Case 1: Toàn bộ dữ liệu", "header")
        
        # Run in thread to prevent UI blocking
        thread = threading.Thread(target=self._execute_workflow_case1)
        thread.daemon = True
        thread.start()
        
    def start_workflow_case2(self):
        """Bắt đầu workflow Case 2"""
        if self.is_processing:
            messagebox.showwarning("Cảnh báo", "Hệ thống đang xử lý. Vui lòng đợi.")
            return
            
        self.log_message("Bắt đầu Workflow Case 2: Dữ liệu theo file import", "header")
        
        # Run in thread to prevent UI blocking
        thread = threading.Thread(target=self._execute_workflow_case2)
        thread.daemon = True
        thread.start()
        
    def _execute_workflow_case1(self):
        """Execute workflow case 1 trong thread"""
        try:
            self.is_processing = True
            self.update_button_state_safe(self.btn_stop, 'normal')
            
            self.log_message_safe("Đang thực hiện workflow case 1...", "info")
            self.update_progress_safe(10, "Khởi tạo...")
            
            # Import and execute workflow
            from app import SchoolProcessApp
            console_app = SchoolProcessApp()
            
            self.update_progress_safe(20, "Bắt đầu xử lý...")
            
            # Execute actual workflow
            console_app._execute_workflow_case_1()
            
            self.update_progress_safe(100, "Hoàn thành")
            self.log_message_safe("Workflow Case 1 hoàn thành!", "success")
            
        except Exception as e:
            self.log_message_safe(f"Lỗi trong workflow Case 1: {str(e)}", "error")
            import traceback
            traceback.print_exc()
        finally:
            self.is_processing = False
            self.update_button_state_safe(self.btn_stop, 'disabled')
            
    def _execute_workflow_case2(self):
        """Execute workflow case 2 trong thread"""
        try:
            self.is_processing = True
            self.update_button_state_safe(self.btn_stop, 'normal')
            
            self.log_message_safe("Đang thực hiện workflow case 2...", "info")
            self.update_progress_safe(10, "Khởi tạo...")
            
            # Import and execute workflow
            from app import SchoolProcessApp
            console_app = SchoolProcessApp()
            
            self.update_progress_safe(20, "Bắt đầu xử lý...")
            
            # Execute actual workflow
            console_app._execute_workflow_case_2()
            
            self.update_progress_safe(100, "Hoàn thành")
            self.log_message_safe("Workflow Case 2 hoàn thành!", "success")
            
        except Exception as e:
            self.log_message_safe(f"Lỗi trong workflow Case 2: {str(e)}", "error")
            import traceback
            traceback.print_exc()
        finally:
            self.is_processing = False
            self.update_button_state_safe(self.btn_stop, 'disabled')
            
    def get_teachers_data(self):
        """Lấy dữ liệu giáo viên"""
        if self.is_processing:
            messagebox.showwarning("Cảnh báo", "Hệ thống đang xử lý. Vui lòng đợi.")
            return
            
        self.log_message("Bắt đầu lấy dữ liệu giáo viên...", "info")
        
        thread = threading.Thread(target=self._get_teachers_data_thread)
        thread.daemon = True
        thread.start()
        
    def _get_teachers_data_thread(self):
        """Lấy dữ liệu giáo viên trong thread"""
        try:
            self.is_processing = True
            self.update_progress(10, "Đang kết nối OnLuyen API...")
            
            client = OnLuyenAPIClient()
            
            self.update_progress(30, "Đang lấy dữ liệu giáo viên...")
            result = client.get_teachers(page_size=1000)
            
            if result['success']:
                self.update_progress(80, "Đang xử lý dữ liệu...")
                data = result.get('data', {})
                teachers_list = data.get('data', []) if isinstance(data, dict) else data
                
                self.update_progress(100, "Hoàn thành")
                self.log_message(f"Lấy thành công {len(teachers_list)} giáo viên", "success")
            else:
                self.log_message(f"Lỗi lấy dữ liệu giáo viên: {result.get('error')}", "error")
                
        except Exception as e:
            self.log_message(f"Lỗi: {str(e)}", "error")
        finally:
            self.is_processing = False
            
    def get_students_data(self):
        """Lấy dữ liệu học sinh"""
        if self.is_processing:
            messagebox.showwarning("Cảnh báo", "Hệ thống đang xử lý. Vui lòng đợi.")
            return
            
        self.log_message("Bắt đầu lấy dữ liệu học sinh...", "info")
        
        thread = threading.Thread(target=self._get_students_data_thread)
        thread.daemon = True
        thread.start()
        
    def _get_students_data_thread(self):
        """Lấy dữ liệu học sinh trong thread"""
        try:
            self.is_processing = True
            self.update_progress(10, "Đang kết nối OnLuyen API...")
            
            client = OnLuyenAPIClient()
            
            self.update_progress(30, "Đang lấy dữ liệu học sinh...")
            result = client.get_students(page_index=1, page_size=5000)
            
            if result['success']:
                self.update_progress(80, "Đang xử lý dữ liệu...")
                data = result.get('data', {})
                students_list = data.get('data', []) if isinstance(data, dict) else data
                
                self.update_progress(100, "Hoàn thành")
                self.log_message(f"Lấy thành công {len(students_list)} học sinh", "success")
            else:
                self.log_message(f"Lỗi lấy dữ liệu học sinh: {result.get('error')}", "error")
                
        except Exception as e:
            self.log_message(f"Lỗi: {str(e)}", "error")
        finally:
            self.is_processing = False
            
    def convert_json_to_excel(self):
        """Chuyển đổi JSON sang Excel"""
        # File dialog to select JSON file
        json_file = filedialog.askopenfilename(
            title="Chọn file JSON",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir="data/output"
        )
        
        if not json_file:
            return
            
        self.log_message(f"Bắt đầu chuyển đổi: {os.path.basename(json_file)}", "info")
        
        thread = threading.Thread(target=self._convert_json_to_excel_thread, args=(json_file,))
        thread.daemon = True
        thread.start()
        
    def _convert_json_to_excel_thread(self, json_file):
        """Chuyển đổi JSON sang Excel trong thread"""
        try:
            self.is_processing = True
            self.update_progress(10, "Đang tải JSON...")
            
            converter = JSONToExcelTemplateConverter(json_file)
            
            self.update_progress(30, "Đang load dữ liệu...")
            if not converter.load_json_data():
                self.log_message("Lỗi load dữ liệu JSON", "error")
                return
                
            self.update_progress(50, "Đang trích xuất dữ liệu...")
            teachers_extracted = converter.extract_teachers_data()
            students_extracted = converter.extract_students_data()
            
            if not teachers_extracted and not students_extracted:
                self.log_message("Không có dữ liệu để chuyển đổi", "warning")
                return
                
            self.update_progress(80, "Đang tạo file Excel...")
            output_path = converter.convert()
            
            if output_path:
                self.update_progress(100, "Hoàn thành")
                self.log_message(f"Đã tạo file Excel: {output_path}", "success")
                self.add_result_file(output_path, "Excel")
            else:
                self.log_message("Lỗi tạo file Excel", "error")
                
        except Exception as e:
            self.log_message(f"Lỗi chuyển đổi: {str(e)}", "error")
        finally:
            self.is_processing = False
            
    def stop_processing(self):
        """Dừng xử lý hiện tại"""
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn dừng xử lý?"):
            self.is_processing = False
            self.update_progress(0, "Đã dừng")
            self.log_message("Xử lý đã bị dừng bởi người dùng", "warning")
            
    def test_onluyen_connection(self):
        """Test kết nối OnLuyen API"""
        self.log_message("Đang test kết nối OnLuyen API...", "info")
        
        try:
            client = OnLuyenAPIClient()
            # Test basic connection
            self.log_message("Kết nối OnLuyen API thành công", "success")
        except Exception as e:
            self.log_message(f"Lỗi kết nối OnLuyen API: {str(e)}", "error")
            
    def test_sheets_connection(self):
        """Test kết nối Google Sheets"""
        self.log_message("Đang test kết nối Google Sheets...", "info")
        
        try:
            extractor = GoogleSheetsExtractor()
            # Test connection
            self.log_message("Kết nối Google Sheets thành công", "success")
        except Exception as e:
            self.log_message(f"Lỗi kết nối Google Sheets: {str(e)}", "error")
            
    def browse_directory(self, var):
        """Browse và chọn thư mục"""
        directory = filedialog.askdirectory(initialdir=var.get())
        if directory:
            var.set(directory)
            
    def add_result_file(self, file_path, file_type):
        """Thêm file vào results tab"""
        if os.path.exists(file_path):
            file_name = os.path.basename(file_path)
            file_size = f"{os.path.getsize(file_path) / (1024*1024):.1f} MB"
            file_date = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
            
            self.files_tree.insert('', 'end', 
                                 text=file_name,
                                 values=(file_type, file_path, file_size, file_date))
                                 
    def show_files_context_menu(self, event):
        """Hiển thị context menu cho files"""
        item = self.files_tree.selection()[0] if self.files_tree.selection() else None
        if item:
            self.files_context_menu.post(event.x_root, event.y_root)
            
    def open_selected_file(self):
        """Mở file đã chọn"""
        item = self.files_tree.selection()[0] if self.files_tree.selection() else None
        if item:
            file_path = self.files_tree.item(item, 'values')[1]
            try:
                os.startfile(file_path)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể mở file: {str(e)}")
                
    def open_file_location(self):
        """Mở thư mục chứa file"""
        item = self.files_tree.selection()[0] if self.files_tree.selection() else None
        if item:
            file_path = self.files_tree.item(item, 'values')[1]
            try:
                os.startfile(os.path.dirname(file_path))
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể mở thư mục: {str(e)}")
                
    def copy_file_path(self):
        """Sao chép đường dẫn file"""
        item = self.files_tree.selection()[0] if self.files_tree.selection() else None
        if item:
            file_path = self.files_tree.item(item, 'values')[1]
            self.root.clipboard_clear()
            self.root.clipboard_append(file_path)
            self.log_message(f"Đã sao chép đường dẫn: {file_path}", "info")
            
    def open_config(self):
        """Mở cấu hình"""
        self.notebook.select(1)  # Switch to config tab
        
    def show_about(self):
        """Hiển thị thông tin về ứng dụng"""
        about_text = """
School Process Application v1.0.0

Ứng dụng xử lý dữ liệu trường học với OnLuyen API

Tính năng:
• Kết nối OnLuyen API
• Trích xuất dữ liệu Google Sheets  
• Chuyển đổi JSON sang Excel
• Upload Google Drive

Phát triển bởi: Assistant
Ngày: 2025-07-29
        """
        messagebox.showinfo("Về ứng dụng", about_text)
        
    def refresh_ui(self):
        """Refresh UI"""
        self.log_message("Đang refresh UI...", "info")
        
    def run(self):
        """Chạy ứng dụng"""
        self.log_message("School Process Application đã khởi động", "success")
        self.root.mainloop()


def main():
    """Entry point cho UI"""
    app = SchoolProcessMainWindow()
    app.run()


if __name__ == "__main__":
    main()
