# build_merge_to_pdf.py
"""
打包 merge_to_pdf.py 为独立的 exe 文件
"""

import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime

def get_script_dir():
    """获取脚本所在目录"""
    return Path(__file__).parent.absolute()

def build_exe():
    """使用 PyInstaller 打包 merge_to_pdf.py"""
    script_dir = get_script_dir()
    script_path = script_dir / "merge_to_pdf.py"
    
    if not script_path.exists():
        print(f"[错误] 找不到文件 {script_path}")
        return False
    
    print("=" * 60)
    print("开始打包 merge_to_pdf.py 为 exe 文件")
    print("=" * 60)
    print(f"脚本路径: {script_path}")
    print()
    
    # PyInstaller 命令参数
    cmd = [
        "pyinstaller",
        "--onefile",  # 打包成单个exe文件
        "--console",  # 显示控制台窗口
        "--name=merge_to_pdf",  # 输出文件名
        "--clean",  # 清理临时文件
        "--noconfirm",  # 不确认，直接覆盖
        "--hidden-import=PIL._tkinter_finder",  # 隐藏导入PIL相关模块
        "--hidden-import=PIL.Image",  # 隐藏导入PIL Image
        "--hidden-import=PIL.ImageDraw",  # 隐藏导入PIL ImageDraw
        "--hidden-import=PIL.ImageFont",  # 隐藏导入PIL ImageFont
        "--collect-all=PIL",  # 收集所有PIL相关文件和数据
        "--hidden-import=pptx",  # 隐藏导入pptx（PPT支持）
        "--hidden-import=pptx.util",  # 隐藏导入pptx.util
        "--hidden-import=lxml",  # pptx依赖lxml
        str(script_path)
    ]
    
    # 添加图标（如果存在）
    icon_path = script_dir / "icon.ico"
    if icon_path.exists():
        cmd.extend(["--icon", str(icon_path)])
        print(f"[OK] 使用图标: {icon_path}")
    
    print(f"执行命令: {' '.join(cmd)}")
    print()
    
    try:
        # 执行 PyInstaller
        result = subprocess.run(cmd, cwd=str(script_dir), check=True, capture_output=False)
        
        # 检查输出文件
        dist_file = script_dir / "dist" / "merge_to_pdf.exe"
        if dist_file.exists():
            print()
            print("=" * 60)
            print("[成功] 打包成功！")
            print("=" * 60)
            print(f"输出文件: {dist_file}")
            print(f"文件大小: {dist_file.stat().st_size / 1024 / 1024:.2f} MB")
            print()
            
            # 尝试重命名文件，添加时间戳
            try:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                new_name = f"merge_to_pdf_{timestamp}.exe"
                new_path = script_dir / "dist" / new_name
                dist_file.rename(new_path)
                print(f"[OK] 已重命名为: {new_name}")
                print(f"完整路径: {new_path}")
            except Exception as e:
                print(f"[警告] 重命名失败（可忽略）: {e}")
                print(f"原文件路径: {dist_file}")
            
            return True
        else:
            print()
            print("[错误] 未找到生成的 exe 文件")
            return False
            
    except subprocess.CalledProcessError as e:
        print()
        print("[错误] 打包失败")
        print(f"错误代码: {e.returncode}")
        return False
    except FileNotFoundError:
        print()
        print("[错误] 未找到 PyInstaller")
        print("   请先安装: pip install pyinstaller")
        return False
    except Exception as e:
        print()
        print(f"[错误] 打包失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主函数"""
    print("正在检查 PyInstaller...")
    
    # 检查 PyInstaller 是否安装
    try:
        import PyInstaller
        print(f"[OK] PyInstaller 版本: {PyInstaller.__version__}")
    except ImportError:
        print("[错误] 未安装 PyInstaller")
        print("   正在尝试安装...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
            print("[OK] PyInstaller 安装成功")
        except Exception as e:
            print(f"[错误] 安装失败: {e}")
            print("   请手动安装: pip install pyinstaller")
            input("按回车键退出...")
            return
    
    print()
    
    # 执行打包
    success = build_exe()
    
    if success:
        print()
        print("=" * 60)
        print("打包完成！")
        print("=" * 60)
    else:
        print()
        print("=" * 60)
        print("打包失败，请检查错误信息")
        print("=" * 60)
    
    try:
        input("按回车键退出...")
    except (EOFError, KeyboardInterrupt):
        pass  # 非交互式环境，忽略输入错误

if __name__ == "__main__":
    main()
