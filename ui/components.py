"""
UI Components for School Process Application
Các component UI tái sử dụng
"""

import tkinter as tk
from tkinter import ttk
from datetime import datetime


class StatusIndicator(ttk.Frame):
    """Component hiển thị trạng thái với màu sắc"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.status_colors = {
            'success': '#4CAF50',
            'error': '#F44336', 
            'warning': '#FF9800',
            'info': '#2196F3',
            'default': '#757575'
        }
        
        self.icon_label = ttk.Label(self, text="●")
        self.icon_label.pack(side='left', padx=(0, 5))
        
        self.text_label = ttk.Label(self, text="")
        self.text_label.pack(side='left')
        
    def set_status(self, text, status='default'):
        """Cập nhật trạng thái"""
        self.text_label.config(text=text)
        color = self.status_colors.get(status, self.status_colors['default'])
        self.icon_label.config(foreground=color)


class ProgressCard(ttk.LabelFrame):
    """Card hiển thị tiến trình với thông tin chi tiết"""
    
    def __init__(self, parent, title="Tiến trình", **kwargs):
        super().__init__(parent, text=title, padding="10", **kwargs)
        
        self.setup_ui()
        
    def setup_ui(self):
        """Thiết lập UI cho progress card"""
        # Task label
        self.task_var = tk.StringVar(value="Sẵn sàng")
        self.task_label = ttk.Label(self, textvariable=self.task_var)
        self.task_label.pack(fill='x', pady=(0, 5))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self, 
                                          variable=self.progress_var,
                                          maximum=100)
        self.progress_bar.pack(fill='x', pady=(0, 5))
        
        # Percentage label
        self.percent_var = tk.StringVar(value="0%")
        self.percent_label = ttk.Label(self, textvariable=self.percent_var)
        self.percent_label.pack()
        
    def update_progress(self, value, task=""):
        """Cập nhật tiến trình"""
        self.progress_var.set(value)
        self.percent_var.set(f"{value:.0f}%")
        
        if task:
            self.task_var.set(task)


class LogViewer(ttk.Frame):
    """Component xem log với màu sắc và filter"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.setup_ui()
        self.setup_tags()
        
    def setup_ui(self):
        """Thiết lập UI"""
        # Control frame
        control_frame = ttk.Frame(self)
        control_frame.pack(fill='x', pady=(0, 5))
        
        # Filter buttons
        ttk.Button(control_frame, text="🗑️ Xóa", 
                  command=self.clear_log).pack(side='left', padx=(0, 5))
        
        ttk.Button(control_frame, text="💾 Lưu", 
                  command=self.save_log).pack(side='left', padx=(0, 5))
        
        # Auto scroll checkbox
        self.auto_scroll_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(control_frame, text="Auto scroll", 
                       variable=self.auto_scroll_var).pack(side='right')
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(self)
        text_frame.pack(fill='both', expand=True)
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)
        
        self.text_widget = tk.Text(text_frame,
                                  wrap='word',
                                  font=('Consolas', 9),
                                  bg='#1e1e1e',
                                  fg='#ffffff',
                                  insertbackground='#ffffff')
        self.text_widget.grid(row=0, column=0, sticky='nsew')
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', 
                                command=self.text_widget.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.text_widget.configure(yscrollcommand=scrollbar.set)
        
    def setup_tags(self):
        """Thiết lập tags cho colored output"""
        tags = {
            "success": {"foreground": "#4CAF50"},
            "error": {"foreground": "#F44336"},
            "warning": {"foreground": "#FF9800"},
            "info": {"foreground": "#2196F3"},
            "header": {"foreground": "#9C27B0", "font": ('Consolas', 9, 'bold')},
            "timestamp": {"foreground": "#757575"}
        }
        
        for tag, config in tags.items():
            self.text_widget.tag_configure(tag, **config)
            
    def add_log(self, message, level="info"):
        """Thêm log message"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Insert timestamp
        self.text_widget.insert(tk.END, f"[{timestamp}] ", "timestamp")
        
        # Insert message with level tag
        self.text_widget.insert(tk.END, f"{message}\n", level)
        
        # Auto scroll if enabled
        if self.auto_scroll_var.get():
            self.text_widget.see(tk.END)
            
    def clear_log(self):
        """Xóa log"""
        self.text_widget.delete(1.0, tk.END)
        
    def save_log(self):
        """Lưu log ra file"""
        from tkinter import filedialog
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.text_widget.get(1.0, tk.END))
                self.add_log(f"Đã lưu log: {file_path}", "success")
            except Exception as e:
                self.add_log(f"Lỗi lưu log: {str(e)}", "error")


class FileList(ttk.Frame):
    """Component hiển thị danh sách files với context menu"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.setup_ui()
        self.setup_context_menu()
        
    def setup_ui(self):
        """Thiết lập UI"""
        # Treeview
        self.tree = ttk.Treeview(self, 
                               columns=('Type', 'Size', 'Date'),
                               show='tree headings')
        self.tree.pack(fill='both', expand=True, side='left')
        
        # Configure columns
        self.tree.heading('#0', text='Tên file')
        self.tree.heading('Type', text='Loại')
        self.tree.heading('Size', text='Kích thước')
        self.tree.heading('Date', text='Ngày tạo')
        
        self.tree.column('#0', width=200)
        self.tree.column('Type', width=80)
        self.tree.column('Size', width=100)
        self.tree.column('Date', width=150)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self, orient='vertical', command=self.tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=scrollbar.set)
        
    def setup_context_menu(self):
        """Thiết lập context menu"""
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Mở file", command=self.open_file)
        self.context_menu.add_command(label="Mở thư mục", command=self.open_folder)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Sao chép đường dẫn", command=self.copy_path)
        self.context_menu.add_command(label="Xóa", command=self.delete_file)
        
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Double-1>", lambda e: self.open_file())
        
    def add_file(self, file_path, file_type=""):
        """Thêm file vào danh sách"""
        import os
        from datetime import datetime
        
        if os.path.exists(file_path):
            file_name = os.path.basename(file_path)
            file_size = f"{os.path.getsize(file_path) / (1024*1024):.1f} MB"
            file_date = datetime.fromtimestamp(
                os.path.getmtime(file_path)
            ).strftime("%Y-%m-%d %H:%M:%S")
            
            # Store file path in tags
            item_id = self.tree.insert('', 'end',
                                     text=file_name,
                                     values=(file_type, file_size, file_date),
                                     tags=(file_path,))
            return item_id
        return None
        
    def get_selected_file(self):
        """Lấy file đang được chọn"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            tags = self.tree.item(item, 'tags')
            if tags:
                return tags[0]  # File path
        return None
        
    def show_context_menu(self, event):
        """Hiển thị context menu"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
            
    def open_file(self):
        """Mở file"""
        file_path = self.get_selected_file()
        if file_path:
            import os
            try:
                os.startfile(file_path)
            except Exception as e:
                print(f"Lỗi mở file: {e}")
                
    def open_folder(self):
        """Mở thư mục chứa file"""
        file_path = self.get_selected_file()
        if file_path:
            import os
            try:
                os.startfile(os.path.dirname(file_path))
            except Exception as e:
                print(f"Lỗi mở thư mục: {e}")
                
    def copy_path(self):
        """Sao chép đường dẫn"""
        file_path = self.get_selected_file()
        if file_path:
            self.clipboard_clear()
            self.clipboard_append(file_path)
            
    def delete_file(self):
        """Xóa file"""
        file_path = self.get_selected_file()
        if file_path:
            from tkinter import messagebox
            import os
            
            if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa file này?"):
                try:
                    os.remove(file_path)
                    # Remove from tree
                    selection = self.tree.selection()
                    if selection:
                        self.tree.delete(selection[0])
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Không thể xóa file: {e}")


class ConfigSection(ttk.LabelFrame):
    """Component section cấu hình"""
    
    def __init__(self, parent, title, **kwargs):
        super().__init__(parent, text=title, padding="10", **kwargs)
        
        self.fields = {}
        self.row_count = 0
        
        self.columnconfigure(1, weight=1)
        
    def add_field(self, label, field_type="entry", **kwargs):
        """Thêm field cấu hình"""
        # Label
        ttk.Label(self, text=f"{label}:").grid(
            row=self.row_count, column=0, sticky='w', pady=(0, 5)
        )
        
        # Field widget
        if field_type == "entry":
            var = tk.StringVar(value=kwargs.get('value', ''))
            widget = ttk.Entry(self, textvariable=var, width=50)
            self.fields[label] = var
            
        elif field_type == "password":
            var = tk.StringVar(value=kwargs.get('value', ''))
            widget = ttk.Entry(self, textvariable=var, show="*", width=50)
            self.fields[label] = var
            
        elif field_type == "checkbox":
            var = tk.BooleanVar(value=kwargs.get('value', False))
            widget = ttk.Checkbutton(self, variable=var, text=kwargs.get('text', ''))
            self.fields[label] = var
            
        elif field_type == "combobox":
            var = tk.StringVar(value=kwargs.get('value', ''))
            widget = ttk.Combobox(self, textvariable=var, 
                                values=kwargs.get('values', []), width=47)
            self.fields[label] = var
            
        widget.grid(row=self.row_count, column=1, sticky='ew', 
                   pady=(0, 5), padx=(10, 0))
        
        # Browse button for file/directory fields
        if kwargs.get('browse'):
            browse_btn = ttk.Button(self, text="...", width=3,
                                  command=lambda: self._browse_file(var, kwargs.get('browse')))
            browse_btn.grid(row=self.row_count, column=2, pady=(0, 5), padx=(5, 0))
            
        self.row_count += 1
        
    def _browse_file(self, var, browse_type):
        """Browse file/directory"""
        from tkinter import filedialog
        
        if browse_type == 'file':
            path = filedialog.askopenfilename()
        elif browse_type == 'directory':
            path = filedialog.askdirectory()
        else:
            return
            
        if path:
            var.set(path)
            
    def get_values(self):
        """Lấy tất cả giá trị"""
        return {label: var.get() for label, var in self.fields.items()}
        
    def set_values(self, values):
        """Thiết lập giá trị"""
        for label, value in values.items():
            if label in self.fields:
                self.fields[label].set(value)


class WorkflowCard(ttk.LabelFrame):
    """Card hiển thị workflow với các bước"""
    
    def __init__(self, parent, title, steps, **kwargs):
        super().__init__(parent, text=title, padding="10", **kwargs)
        
        self.steps = steps
        self.step_widgets = {}
        
        self.setup_ui()
        
    def setup_ui(self):
        """Thiết lập UI"""
        for i, step in enumerate(self.steps):
            step_frame = ttk.Frame(self)
            step_frame.pack(fill='x', pady=(0, 5))
            
            # Step icon/number
            icon_label = ttk.Label(step_frame, text=f"{i+1}️⃣", font=('Segoe UI', 12))
            icon_label.pack(side='left', padx=(0, 10))
            
            # Step text
            text_label = ttk.Label(step_frame, text=step, font=('Segoe UI', 10))
            text_label.pack(side='left', fill='x', expand=True)
            
            # Status indicator
            status_label = ttk.Label(step_frame, text="⏳", font=('Segoe UI', 12))
            status_label.pack(side='right')
            
            self.step_widgets[i] = {
                'frame': step_frame,
                'text': text_label,
                'status': status_label
            }
            
    def update_step(self, step_index, status='pending'):
        """Cập nhật trạng thái bước"""
        if step_index in self.step_widgets:
            status_icons = {
                'pending': '⏳',
                'running': '🔄',
                'success': '✅',
                'error': '❌',
                'warning': '⚠️'
            }
            
            icon = status_icons.get(status, '⏳')
            self.step_widgets[step_index]['status'].config(text=icon)
            
    def reset_steps(self):
        """Reset tất cả bước về pending"""
        for i in range(len(self.steps)):
            self.update_step(i, 'pending')
