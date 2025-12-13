@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
echo ========================================
echo 高级优化打包（最小体积）
echo ========================================
echo.
echo 此脚本将进行深度优化，进一步减小exe体积
echo 注意：可能需要更长的打包时间
echo.
pause

REM 检查依赖
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo 正在安装PyInstaller...
    pip install pyinstaller
)

REM 清理
echo 清理构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
for /d /r . %%d in (__pycache__) do @if exist "%%d" rmdir /s /q "%%d"

REM 使用命令行方式打包，添加更多优化选项
echo.
echo 开始高级优化打包...
echo.

pyinstaller ^
    --name=route_system_baidu ^
    --onefile ^
    --windowed ^
    --clean ^
    --noconfirm ^
    --optimize=2 ^
    --strip ^
    --noupx ^
    --exclude-module=matplotlib ^
    --exclude-module=scipy ^
    --exclude-module=tkinter ^
    --exclude-module=PyQt5 ^
    --exclude-module=PyQt6 ^
    --exclude-module=IPython ^
    --exclude-module=jupyter ^
    --exclude-module=notebook ^
    --exclude-module=pytest ^
    --exclude-module=unittest ^
    --exclude-module=sphinx ^
    --exclude-module=pydoc ^
    --add-data "templates;templates" ^
    --add-data "static;static" ^
    --hidden-import=pandas._libs.tslibs.timedeltas ^
    --hidden-import=pandas._libs.tslibs.nattype ^
    --hidden-import=pandas._libs.tslibs.np_datetime ^
    --hidden-import=pandas._libs.skiplist ^
    --hidden-import=numpy.random ^
    --hidden-import=numpy.random.bit_generator ^
    --hidden-import=numpy.random._generator ^
    --hidden-import=numpy.random._bounded_integers ^
    --hidden-import=openpyxl ^
    --hidden-import=flask ^
    --hidden-import=werkzeug ^
    --hidden-import=jinja2 ^
    --hidden-import=requests ^
    app.py

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
    echo ========================================
    echo.
    
    if exist dist\route_system_baidu.exe (
        for %%A in (dist\route_system_baidu.exe) do (
            set /a sizeMB=%%~zA/1024/1024
            echo EXE文件位置: dist\route_system_baidu.exe
            echo EXE文件大小: !sizeMB! MB
        )
    )
    
    echo.
    pause
)

