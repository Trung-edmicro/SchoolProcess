@echo off
echo =========================================
echo    School Process - UI Application
echo =========================================
echo.

cd /d "%~dp0"

echo Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.8+ and add it to PATH
    pause
    exit /b 1
)

echo Starting School Process UI...
echo.

python main_ui.py

echo.
echo Application closed.
pause
