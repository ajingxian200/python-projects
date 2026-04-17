@echo off
echo ========================================
echo   Excel 对比分析工具 - 打包脚本
echo ========================================
echo.

echo [1/2] 安装依赖...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo 依赖安装失败！
    pause
    exit /b 1
)

echo.
echo [2/2] 打包为 exe...
pyinstaller --onefile --windowed --name "Excel分析工具" --clean app.py
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
