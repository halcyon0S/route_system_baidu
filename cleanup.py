#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
工作目录清理脚本 - 删除非必要文件
"""

import os
import shutil
from pathlib import Path

def get_size_mb(path):
    """获取文件或目录大小（MB）"""
    if os.path.isfile(path):
        return os.path.getsize(path) / (1024 * 1024)
    elif os.path.isdir(path):
        total = 0
        for dirpath, dirnames, filenames in os.walk(path):
            for filename in filenames:
                filepath = os.path.join(dirpath, filename)
                if os.path.isfile(filepath):
                    total += os.path.getsize(filepath)
        return total / (1024 * 1024)
    return 0

def safe_remove(path, description="", skip_if_in_use=True):
    """安全删除文件或目录"""
    if not os.path.exists(path):
        return 0, 0
    
    size_mb = get_size_mb(path)
    try:
        if os.path.isdir(path):
            # 对于目录，先尝试删除内容
            if skip_if_in_use:
                # 尝试删除目录内容，如果失败则跳过
                try:
                    shutil.rmtree(path, ignore_errors=False)
                except PermissionError:
                    print(f"  [SKIP] 跳过 (可能正在使用): {path} {description}")
                    return 0, 0
            else:
                shutil.rmtree(path)
        else:
            if skip_if_in_use:
                try:
                    os.remove(path)
                except PermissionError:
                    print(f"  [SKIP] 跳过 (可能正在使用): {path} {description}")
                    return 0, 0
            else:
                os.remove(path)
        print(f"  [OK] 已删除: {path} {description} ({size_mb:.2f} MB)")
        return 1, size_mb
    except Exception as e:
        error_msg = str(e)
        if "Permission denied" in error_msg or "拒绝访问" in error_msg or "WinError 5" in error_msg:
            print(f"  [SKIP] 跳过 (权限不足或文件正在使用): {path}")
        else:
            print(f"  [FAIL] 删除失败: {path} - {e}")
        return 0, 0

def main():
    """主函数"""
    print("=" * 60)
    print("工作目录清理脚本")
    print("=" * 60)
    print()
    
    base_dir = Path(".")
    removed_count = 0
    total_size_mb = 0
    
    # 1. 删除构建产物目录
    print("[1] 清理构建产物...")
    # build目录可以完全删除
    count, size = safe_remove("build", "(构建目录)")
    removed_count += count
    total_size_mb += size
    
    # dist目录可能包含用户需要的exe文件，提示用户手动处理
    if os.path.exists("dist"):
        dist_size = get_size_mb("dist")
        print(f"  [INFO] dist目录存在 ({dist_size:.2f} MB)，如需删除请手动处理")
        print(f"  [INFO] 建议：如果不需要exe文件，可以手动删除 dist 目录")
    
    # 2. 删除日志目录
    print("\n[2] 清理日志文件...")
    if os.path.exists("logs"):
        count, size = safe_remove("logs", "(日志目录)")
        removed_count += count
        total_size_mb += size
    else:
        print("  (logs目录不存在)")
    
    # 3. 删除测试Excel文件（保留模板文件）
    print("\n[3] 清理测试Excel文件...")
    test_excel_files = [
        "11.xlsx",
        "包涵.xlsx",
        "导入网组网点表（南京、测试2人）1.xlsx",
        "导入网组网点表（测试表，2人） - 副本 (2) - 副本.xlsx",
        "导入网组网点表（测试表，2人） - 副本 (2).xlsx",
        "导入网组网点表（测试表，2人） - 副本.xlsx",
        "导入网组网点表（测试表，2人）.xlsx",
    ]
    # 保留的模板文件
    keep_files = ["地点模板.xlsx", "地点模板工号版.xlsx"]
    
    for filename in test_excel_files:
        if filename not in keep_files:
            # 使用Path对象处理中文文件名编码问题
            filepath = Path(filename)
            if filepath.exists():
                try:
                    count, size = safe_remove(str(filepath), f"(测试文件)")
                    removed_count += count
                    total_size_mb += size
                except Exception as e:
                    print(f"  [SKIP] 跳过文件 (编码问题): {filename}")
    
    # 4. 删除输出目录
    print("\n[4] 清理输出目录...")
    output_dirs = [
        "合并PDF",
        "合并PPT",
        "网点图",
        "网组网点路线图",
    ]
    for dir_name in output_dirs:
        count, size = safe_remove(dir_name, f"(输出目录)")
        removed_count += count
        total_size_mb += size
    
    # 5. 删除生成的资源文件
    print("\n[5] 清理生成的资源文件...")
    if os.path.exists("遮罩图片"):
        count, size = safe_remove("遮罩图片", "(生成资源)")
        removed_count += count
        total_size_mb += size
    else:
        print("  (遮罩图片目录不存在)")
    
    # 6. 清理Python缓存（如果存在）
    print("\n[6] 清理Python缓存...")
    cache_items = ["__pycache__"]
    for item in cache_items:
        if os.path.exists(item):
            count, size = safe_remove(item, "(Python缓存)")
            removed_count += count
            total_size_mb += size
    
    # 总结
    print()
    print("=" * 60)
    print("清理完成")
    print("=" * 60)
    print(f"删除项目数: {removed_count}")
    print(f"释放空间: {total_size_mb:.2f} MB")
    print()
    print("保留的必要文件/目录:")
    print("  - app.py, jietu.py, merge_to_pdf.py, generate_mask_images.py")
    print("  - templates/, static/")
    print("  - build.py, build.bat, build_console.bat, build_merge_to_pdf.bat, build_merge_to_pdf.py")
    print("  - route_system_baidu.spec")
    print("  - requirements.txt, README_打包说明.md")
    print("  - config-custom.js")
    print("  - 地点模板.xlsx, 地点模板工号版.xlsx")
    print("  - xzqh.html")
    print()
    print("注意:")
    print("  - 下次打包时会重新生成 build/ 和 dist/ 目录")
    print("  - 运行程序后会自动创建输出目录")
    print("  - 如需遮罩图片，运行 generate_mask_images.py 重新生成")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n用户中断操作")
    except Exception as e:
        print(f"\n\n发生错误: {e}")
    finally:
        pass  # 移除交互式输入，允许在非交互环境中运行

