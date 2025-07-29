@echo off
echo 🏫 School Process - Khởi động ứng dụng...

REM Kiểm tra Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python không được tìm thấy!
    echo Vui lòng cài đặt Python trước khi chạy ứng dụng.
    pause
    exit /b 1
)

REM Kiểm tra virtual environment
if exist ".venv\Scripts\python.exe" (
    echo ✅ Sử dụng virtual environment
    .venv\Scripts\python.exe main_ui.py
) else (
    echo ⚠️ Không tìm thấy virtual environment, sử dụng Python system
    python main_ui.py
)

echo 📱 Ứng dụng đã đóng.
pause
