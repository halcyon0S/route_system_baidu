@echo off
chcp 65001 >nul
echo ========================================
echo 正在打包为 EXE 文件...
echo ========================================
echo.

REM 检查是否安装了 PyInstaller
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [错误] 未安装 PyInstaller
    echo 正在安装 PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo [错误] PyInstaller 安装失败
        pause
        exit /b 1
    )
)

REM 清理之前的构建文件
echo [1/3] 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__

REM 执行打包
echo [2/3] 正在打包...
pyinstaller route_system_baidu.spec --clean --noconfirm

if errorlevel 1 (
    echo.
    echo [错误] 打包失败！
    pause
    exit /b 1
)

REM 检查输出文件
echo [3/3] 检查输出文件...
if exist "dist\route_system_baidu.exe" (
    echo.
    echo ========================================
    echo ✓ 打包成功！
    echo ========================================
    echo.
    echo EXE 文件位置: dist\route_system_baidu.exe
    echo.
    dir "dist\route_system_baidu.exe" | findstr "route_system_baidu.exe"
    echo.
    echo 提示：双击 dist\route_system_baidu.exe 即可运行
    echo.
) else (
    echo.
    echo [错误] 未找到输出文件 dist\route_system_baidu.exe
    pause
    exit /b 1
)

pause
