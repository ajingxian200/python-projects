@echo off
echo ========================================
echo   CSV 转 Excel 工具 - 打包脚本
echo ========================================
echo.

echo [1/3] 安装依赖...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo 依赖安装失败！
    pause
    exit /b 1
)

echo.
echo [2/3] 安装 PyInstaller...
pip install pyinstaller
if %errorlevel% neq 0 (
    echo PyInstaller 安装失败！
    pause
    exit /b 1
)

echo.
echo [3/3] 打包为 exe...
pyinstaller --onefile --windowed --name "CSV转Excel工具" --clean app.py
if %errorlevel% neq 0 (
    echo 打包失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo   打包完成！exe 文件在 dist 目录下
echo ========================================
pause
