@echo off
chcp 65001 >nul
echo ========================================
echo 打包 merge_to_pdf.py 为 exe 文件
echo ========================================
echo.

python build_merge_to_pdf.py

pause
