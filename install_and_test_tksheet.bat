@echo off
echo ============================================================
echo     CACHING VA TEST TKSHEET CHO SCHOOL PROCESS
echo ============================================================
echo.

echo 🔧 Kich hoat virtual environment...
if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
    echo ✅ Virtual environment kich hoat thanh cong
) else (
    echo ⚠️ Khong tim thay virtual environment
    echo Tao virtual environment moi...
    python -m venv .venv
    call .venv\Scripts\activate.bat
    echo ✅ Virtual environment moi da duoc tao
)

echo.
echo 📦 Cai dat dependencies...
pip install --upgrade pip
pip install -r requirements.txt

echo.
echo 🧪 Kiem tra tksheet...
python -c "import tksheet; print('✅ Tksheet version:', tksheet.__version__)" 2>nul
if errorlevel 1 (
    echo ❌ Tksheet chua duoc cai dat dung cach
    echo Cai dat tksheet rieng biet...
    pip install tksheet>=7.5.0
)

echo.
echo 🚀 Chay demo Tksheet...
python test_tksheet_demo.py

echo.
echo 🎯 Neu demo chay thanh cong, ban co the chay ung dung chinh:
echo python main_ui.py
echo.
pause