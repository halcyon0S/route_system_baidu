"""
Edge 远程调试模式测试
==================

工作原理：
1. 手动启动 Edge 浏览器（带远程调试参数）
2. Selenium 连接到已运行的浏览器实例
3. 优点：浏览器窗口由用户控制，更稳定

使用步骤：
Step 1: 运行此脚本（会自动启动 Edge）
Step 2: Selenium 会自动连接并控制浏览器
"""

import subprocess
import time
import os
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options


def start_edge_with_remote_debugging(debug_port=9222, user_data_dir=None):
    """
    启动带远程调试的 Edge 浏览器
    
    Args:
        debug_port: 远程调试端口
        user_data_dir: 用户数据目录（None 则使用临时目录）
    
    Returns:
        subprocess.Popen 对象
    """
    # Edge 浏览器路径（常见位置）
    edge_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    
    edge_path = None
    for path in edge_paths:
        if os.path.exists(path):
            edge_path = path
            break
    
    if not edge_path:
        raise FileNotFoundError("未找到 Edge 浏览器，请检查安装路径")
    
    # 用户数据目录（避免与正常使用的 Edge 冲突）
    if user_data_dir is None:
        user_data_dir = os.path.join(os.getenv('TEMP'), 'EdgeDebugProfile')
    
    # 启动命令
    cmd = [
        edge_path,
        f'--remote-debugging-port={debug_port}',
        f'--user-data-dir={user_data_dir}',
        '--no-first-run',
        '--no-default-browser-check',
        '--disable-extensions',
    ]
    
    print(f"[启动] Edge 浏览器路径: {edge_path}")
    print(f"[启动] 远程调试端口: {debug_port}")
    print(f"[启动] 用户数据目录: {user_data_dir}")
    print(f"[启动] 执行命令...")
    
    # 启动进程
    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
    )
    
    print(f"[启动] ✓ Edge 已启动 (PID: {process.pid})")
    print(f"[启动] 等待浏览器初始化...")
    time.sleep(3)  # 等待浏览器完全启动
    
    return process


def connect_to_remote_edge(debug_port=9222):
    """
    连接到远程调试模式的 Edge 浏览器
    
    Args:
        debug_port: 远程调试端口
    
    Returns:
        WebDriver 实例
    """
    print(f"\n[连接] 正在连接到远程 Edge (端口: {debug_port})...")
    
    # 配置选项
    options = Options()
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
    
    # 创建驱动（不需要启动新浏览器）
    service = Service()
    driver = webdriver.Edge(service=service, options=options)
    
    print(f"[连接] ✓ 已连接到 Edge 浏览器")
    print(f"[连接] 当前 URL: {driver.current_url}")
    
    return driver


def test_basic_operations(driver):
    """测试基本操作"""
    print(f"\n{'='*60}")
    print("开始测试基本操作")
    print(f"{'='*60}")
    
    # 测试 1: 访问百度
    print("\n[测试1] 访问百度...")
    driver.get("https://www.baidu.com")
    time.sleep(2)
    print(f"  ✓ 当前 URL: {driver.current_url}")
    print(f"  ✓ 页面标题: {driver.title}")
    
    # 测试 2: 访问本地服务器
    print("\n[测试2] 访问本地服务器...")
    local_url = "http://127.0.0.1:5005"
    print(f"  → 尝试访问: {local_url}")
    
    try:
        driver.get(local_url)
        time.sleep(2)
        print(f"  ✓ 当前 URL: {driver.current_url}")
        print(f"  ✓ 页面标题: {driver.title}")
    except Exception as e:
        print(f"  ⚠ 访问失败（服务器可能未运行）: {e}")
    
    # 测试 3: 获取窗口信息
    print("\n[测试3] 获取窗口信息...")
    print(f"  ✓ 窗口句柄: {driver.current_window_handle}")
    print(f"  ✓ 窗口大小: {driver.get_window_size()}")
    print(f"  ✓ 窗口位置: {driver.get_window_position()}")
    
    # 测试 4: 截图
    print("\n[测试4] 测试截图功能...")
    screenshot_path = "test_screenshot.png"
    driver.save_screenshot(screenshot_path)
    if os.path.exists(screenshot_path):
        print(f"  ✓ 截图已保存: {screenshot_path}")
    else:
        print(f"  ✗ 截图保存失败")
    
    print(f"\n{'='*60}")
    print("测试完成！")
    print(f"{'='*60}")


def main():
    """主函数"""
    print("=" * 60)
    print("Edge 远程调试模式测试")
    print("=" * 60)
    
    debug_port = 9222
    edge_process = None
    driver = None
    
    try:
        # Step 1: 启动 Edge（远程调试模式）
        edge_process = start_edge_with_remote_debugging(debug_port)
        
        # Step 2: 连接到 Edge
        driver = connect_to_remote_edge(debug_port)
        
        # Step 3: 执行测试
        test_basic_operations(driver)
        
        # Step 4: 等待用户操作
        print("\n" + "=" * 60)
        print("浏览器将保持打开状态")
        print("你可以手动操作浏览器，Selenium 保持连接")
        print("按 Enter 键关闭测试...")
        print("=" * 60)
        input()
        
    except FileNotFoundError as e:
        print(f"\n❌ 错误: {e}")
    except Exception as e:
        print(f"\n❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # 清理
        print("\n[清理] 正在清理资源...")
        
        if driver:
            try:
                print("[清理] 断开 Selenium 连接...")
                driver.quit()
                print("[清理] ✓ Selenium 已断开")
            except:
                pass
        
        if edge_process:
            try:
                print("[清理] 关闭 Edge 浏览器...")
                edge_process.terminate()
                edge_process.wait(timeout=5)
                print("[清理] ✓ Edge 已关闭")
            except:
                print("[清理] ⚠ Edge 进程可能仍在运行")
        
        print("[清理] 完成")


if __name__ == "__main__":
    main()
