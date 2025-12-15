#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
打包脚本（带控制台窗口版本）- 将应用打包为EXE文件并添加时间戳
"""

import os
import sys
import subprocess
import shutil
import re
import logging
from datetime import datetime

def setup_logging():
    """设置日志记录"""
    # 创建logs目录
    if not os.path.exists("logs"):
        os.makedirs("logs")
    
    # 生成日志文件名（带时间戳）
    log_filename = os.path.join("logs", f"build_console_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    
    # 配置日志格式
    log_format = '%(asctime)s - %(levelname)s - %(message)s'
    date_format = '%Y-%m-%d %H:%M:%S'
    
    # 配置日志记录器
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        datefmt=date_format,
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info(f"日志文件: {log_filename}")
    return logger, log_filename

def check_pyinstaller(logger):
    """检查PyInstaller是否已安装"""
    try:
        import PyInstaller
        logger.info("PyInstaller 已安装")
        return True
    except ImportError:
        logger.warning("未安装 PyInstaller")
        logger.info("正在安装 PyInstaller...")
        result = subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], 
                              capture_output=True, text=True)
        if result.returncode != 0:
            logger.error("PyInstaller 安装失败")
            logger.error(result.stderr)
            return False
        logger.info("PyInstaller 安装成功")
        return True

def create_console_spec(logger):
    """创建控制台版本的spec文件"""
    logger.info("[0/3] 修改配置为控制台模式...")
    spec_file = "route_system_baidu.spec"
    console_spec_file = "route_system_baidu_console.spec"
    
    if not os.path.exists(spec_file):
        logger.error(f"未找到 {spec_file}")
        return False
    
    # 读取原spec文件
    with open(spec_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 替换 console=False 为 console=True
    content = re.sub(r'console=False', 'console=True', content)
    
    # 写入新的spec文件
    with open(console_spec_file, 'w', encoding='utf-8') as f:
        f.write(content)
    
    logger.info(f"  已创建: {console_spec_file}")
    return console_spec_file

def clean_build_files(logger):
    """清理旧的构建文件"""
    logger.info("[1/3] 清理旧的构建文件...")
    dirs_to_remove = ['build']
    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            try:
                shutil.rmtree(dir_name)
                logger.info(f"  已删除: {dir_name}")
            except Exception as e:
                logger.warning(f"  无法删除 {dir_name}: {e}")
    
    # 删除旧的exe文件
    old_exe = os.path.join("dist", "route_system_baidu.exe")
    if os.path.exists(old_exe):
        try:
            os.remove(old_exe)
            logger.info(f"  已删除: {old_exe}")
        except Exception as e:
            logger.warning(f"  无法删除 {old_exe}: {e}")

def build_exe(console_spec_file, logger):
    """执行打包"""
    logger.info("[2/3] 正在打包...")
    logger.info(f"使用配置文件: {console_spec_file}")
    logger.info("执行 PyInstaller 打包命令...")
    
    # 执行打包命令，实时输出
    process = subprocess.Popen([
        sys.executable, "-m", "PyInstaller", 
        console_spec_file, 
        "--clean", 
        "--noconfirm"
    ], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, universal_newlines=True)
    
    # 实时读取输出并记录
    output_lines = []
    for line in process.stdout:
        line = line.rstrip()
        if line:
            logger.info(f"  {line}")
            output_lines.append(line)
    
    process.wait()
    
    if process.returncode != 0:
        logger.error("打包失败！")
        logger.error("\n".join(output_lines[-50:]))  # 只记录最后50行错误信息
        return False
    
    logger.info("打包完成")
    return True

def rename_exe_with_timestamp(logger):
    """重命名exe文件，添加时间戳"""
    logger.info("[3/3] 检查输出文件并添加时间戳...")
    
    old_exe_path = os.path.join("dist", "route_system_baidu.exe")
    if not os.path.exists(old_exe_path):
        logger.error(f"未找到输出文件 {old_exe_path}")
        return None
    
    logger.info(f"找到输出文件: {old_exe_path}")
    
    # 生成时间戳：yyyyMMddHHmm
    timestamp = datetime.now().strftime("%Y%m%d%H%M")
    new_exe_name = f"route_system_{timestamp}.exe"
    new_exe_path = os.path.join("dist", new_exe_name)
    
    logger.info(f"生成时间戳: {timestamp}")
    logger.info(f"新文件名: {new_exe_name}")
    
    # 如果目标文件已存在，先删除
    if os.path.exists(new_exe_path):
        try:
            os.remove(new_exe_path)
            logger.info(f"  已删除已存在的文件: {new_exe_name}")
        except Exception as e:
            logger.warning(f"  无法删除已存在的文件: {e}")
    
    # 重命名文件
    try:
        os.rename(old_exe_path, new_exe_path)
        file_size = os.path.getsize(new_exe_path)
        file_size_mb = file_size / (1024 * 1024)
        logger.info(f"  成功重命名为: {new_exe_name}")
        logger.info(f"  文件大小: {file_size_mb:.2f} MB")
        return new_exe_path
    except Exception as e:
        logger.warning(f"重命名失败: {e}")
        return old_exe_path

def main():
    """主函数"""
    # 设置日志
    logger, log_filename = setup_logging()
    
    logger.info("=" * 40)
    logger.info("正在打包为 EXE 文件（带控制台窗口）...")
    logger.info("=" * 40)
    logger.info("")
    
    start_time = datetime.now()
    logger.info(f"开始时间: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("")
    
    try:
        # 检查PyInstaller
        if not check_pyinstaller(logger):
            logger.error("PyInstaller 检查失败，退出")
            input("按回车键退出...")
            sys.exit(1)
        logger.info("")
        
        # 创建控制台版本的spec文件
        console_spec_file = create_console_spec(logger)
        if not console_spec_file:
            logger.error("创建控制台spec文件失败，退出")
            input("按回车键退出...")
            sys.exit(1)
        logger.info("")
        
        # 清理构建文件
        clean_build_files(logger)
        logger.info("")
        
        # 执行打包
        if not build_exe(console_spec_file, logger):
            # 清理临时文件
            if os.path.exists(console_spec_file):
                os.remove(console_spec_file)
            logger.error("打包失败，退出")
            input("按回车键退出...")
            sys.exit(1)
        logger.info("")
        
        # 清理临时spec文件
        if os.path.exists(console_spec_file):
            try:
                os.remove(console_spec_file)
                logger.info(f"已清理临时文件: {console_spec_file}")
            except Exception as e:
                logger.warning(f"无法删除临时文件: {e}")
        logger.info("")
        
        # 重命名exe文件
        exe_path = rename_exe_with_timestamp(logger)
        logger.info("")
        
        end_time = datetime.now()
        duration = (end_time - start_time).total_seconds()
        
        if exe_path:
            logger.info("=" * 40)
            logger.info("✓ 打包成功！")
            logger.info("=" * 40)
            logger.info("")
            logger.info(f"EXE 文件位置: {exe_path}")
            logger.info(f"原文件名: route_system_baidu.exe")
            if "route_system_" in os.path.basename(exe_path):
                logger.info(f"新文件名: {os.path.basename(exe_path)}")
            else:
                logger.info(f"文件名: {os.path.basename(exe_path)} (未添加时间戳)")
            logger.info("注意：此版本会显示控制台窗口，方便查看日志")
            logger.info("")
            
            # 显示文件信息
            if os.path.exists(exe_path):
                file_size = os.path.getsize(exe_path)
                file_size_mb = file_size / (1024 * 1024)
                logger.info(f"文件大小: {file_size_mb:.2f} MB")
                logger.info("")
            
            logger.info(f"结束时间: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
            logger.info(f"总耗时: {duration:.2f} 秒")
            logger.info(f"日志文件: {log_filename}")
            logger.info("")
            logger.info(f"提示：双击 {exe_path} 即可运行")
            logger.info("")
        else:
            logger.error("打包过程出现问题")
            input("按回车键退出...")
            sys.exit(1)
    
    except Exception as e:
        logger.exception(f"发生未预期的错误: {e}")
        input("按回车键退出...")
        sys.exit(1)
    
    input("按回车键退出...")

if __name__ == "__main__":
    main()
