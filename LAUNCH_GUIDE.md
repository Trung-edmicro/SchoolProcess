# 🏫 School Process - Hướng dẫn chạy ứng dụng

## ✅ Đã khắc phục vấn đề UI chớp tắt

### 🚀 Cách chạy ứng dụng:

#### **Phương án 1: Đơn giản nhất (Khuyến nghị)**
```bash
.\start_app.bat
```
- Tự động kiểm tra Python
- Tự động chọn virtual environment hoặc Python system
- Hiện thông báo lỗi nếu có

#### **Phương án 2: Launcher đơn giản**
```bash
python simple_launcher.py
```
hoặc với virtual environment:
```bash
C:/app/SchoolProcess/.venv/Scripts/python.exe simple_launcher.py
```

#### **Phương án 3: Launcher với splash screen**
```bash
python main_ui.py
```
hoặc với virtual environment:
```bash
C:/app/SchoolProcess/.venv/Scripts/python.exe main_ui.py
```

### 🔧 Nếu UI vẫn chớp tắt:

1. **Chạy từ Command Prompt thay vì double-click file**
   - Mở Command Prompt (cmd)
   - `cd C:\app\SchoolProcess`
   - `.\start_app.bat`

2. **Kiểm tra dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Debug với launcher đơn giản**
   ```bash
   python simple_launcher.py
   ```

### 📋 Thông tin debug:

- ✅ **simple_launcher.py**: Launcher đơn giản không có splash screen
- ✅ **main_ui.py**: Launcher với splash screen (đã sửa threading)
- ✅ **debug_main_window.py**: UI version đơn giản để debug
- ✅ **start_app.bat**: Batch file tự động chọn Python

### 🎯 Kết quả mong đợi:

Khi chạy thành công, bạn sẽ thấy:
1. **Console logs** hiển thị quá trình khởi động
2. **Splash screen** (nếu dùng main_ui.py) 
3. **Main window** với giao diện Material Design
4. **Menu tabs**: Workflow, Logs, Config
5. **Status bar** ở cuối cửa sổ

### ⚠️ Lưu ý:

- Đừng double-click file .py, hãy chạy từ Command Prompt
- Nếu UI vẫn tắt, kiểm tra console để xem lỗi cụ thể
- Virtual environment được khuyến nghị để tránh conflict dependencies
