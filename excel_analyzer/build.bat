@echo off
echo ========================================
echo   CSV to Excel Tool - Build Script
echo ========================================
echo.

echo [1/3] Installing dependencies...
pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple
if %errorlevel% neq 0 (
    echo Dependencies install failed!
    pause
    exit /b 1
)

echo.
echo [2/3] Installing PyInstaller...
pip install pyinstaller
if %errorlevel% neq 0 (
    echo PyInstaller install failed!
    pause
    exit /b 1
)

echo.
echo [3/3] Building exe...
pyinstaller --onefile --windowed --name "CsvToExcel" --clean app.py
if %errorlevel% neq 0 (
    echo Build failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Done! exe is in the dist folder
echo ========================================
pause
