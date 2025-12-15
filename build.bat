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

REM 检查输出文件并重命名（添加时间戳）
echo [3/3] 检查输出文件并添加时间戳...
if exist "dist\route_system_baidu.exe" (
    REM 使用PowerShell完成重命名，时间戳写入文件
    powershell -NoProfile -Command "$ts = (Get-Date).ToString('yyyyMMddHHmm'); $oldFile = 'dist\route_system_baidu.exe'; $newFile = 'dist\route_system_' + $ts + '.exe'; if (Test-Path $oldFile) { if (Test-Path $newFile) { Remove-Item $newFile -Force }; Rename-Item $oldFile $newFile; Write-Host ('成功重命名为: route_system_' + $ts + '.exe'); [System.IO.File]::WriteAllText('timestamp_result.tmp', $ts, [System.Text.Encoding]::ASCII) } else { [System.IO.File]::WriteAllText('timestamp_result.tmp', '', [System.Text.Encoding]::ASCII); exit 1 }" 2>nul
    set /p timestamp=<timestamp_result.tmp
    del timestamp_result.tmp >nul 2>&1
    
    REM 检查重命名是否成功
    if "%timestamp%"=="" (
        echo [警告] 重命名可能失败，检查文件...
        set newfilename=route_system_baidu.exe
        set newfilepath=dist\%newfilename%
    ) else (
        REM 去除可能的换行符
        set timestamp=%timestamp: =%
        set newfilename=route_system_%timestamp%.exe
        set newfilepath=dist\%newfilename%
    )
    
    echo.
    echo ========================================
    echo ✓ 打包成功！
    echo ========================================
    echo.
    echo EXE 文件位置: %newfilepath%
    echo 原文件名: route_system_baidu.exe
    if defined timestamp (
        echo 新文件名: %newfilename%
    ) else (
        echo 文件名: %newfilename% (未添加时间戳)
    )
    echo.
    if exist "%newfilepath%" (
        dir "%newfilepath%" | findstr /i "%newfilename%"
    ) else (
        dir "dist\route_system_baidu.exe" | findstr /i "route_system_baidu"
    )
    echo.
    echo 提示：双击 %newfilepath% 即可运行
    echo.
) else (
    echo.
    echo [错误] 未找到输出文件 dist\route_system_baidu.exe
    pause
    exit /b 1
)

pause
