@echo off
chcp 65001 >nul
echo ========================================
echo 正在打包为EXE文件...
echo ========================================
echo.

REM 检查是否安装了PyInstaller
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo 正在安装PyInstaller...
    pip install pyinstaller
)

REM 执行打包
echo 开始打包...
pyinstaller build_exe.spec --clean

if errorlevel 1 (
    echo.
    echo ========================================
    echo 打包失败！请检查错误信息。
    echo ========================================
    pause
    exit /b 1
) else (
    echo.
    echo ========================================
    echo 打包成功！
    echo EXE文件位置: dist\route_system_baidu.exe
    echo ========================================
    echo.
    echo 提示：
    echo 1. 首次运行可能需要几秒钟启动
    echo 2. 程序会在 http://127.0.0.1:5004 启动
    echo 3. 请在浏览器中打开该地址使用
    echo.
    pause
)

