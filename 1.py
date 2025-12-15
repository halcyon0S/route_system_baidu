"""
app.py 启动器（使用 Edge 远程调试模式）
=====================================

使用方法：
1. 运行此脚本代替直接运行 app.py
2. 脚本会自动：
   - 启动 Edge（远程调试模式）
   - 启动 Flask 服务器
   - 连接 Selenium 到 Edge
   - 打开你的应用
"""

import subprocess
import time
import os
import sys
import signal
import atexit
from threading import Thread


class EdgeDebugLauncher:
    """Edge 远程调试启动器"""
    
    def __init__(self, debug_port=9222):
        self.debug_port = debug_port
        self.edge_process = None
        self.flask_process = None
        self.user_data_dir = os.path.join(os.getenv('TEMP'), 'EdgeDebugProfile_App')
    
    def find_edge_path(self):
        """查找 Edge 浏览器路径"""
        edge_paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]
        
        for path in edge_paths:
            if os.path.exists(path):
                return path
        
        raise FileNotFoundError("未找到 Edge 浏览器")
    
    def start_edge(self):
        """启动 Edge 浏览器（远程调试模式）"""
        edge_path = self.find_edge_path()
        
        cmd = [
            edge_path,
            f'--remote-debugging-port={self.debug_port}',
            f'--user-data-dir={self.user_data_dir}',
            '--no-first-run',
            '--no-default-browser-check',
            '--start-maximized',
            'about:blank',  # 先打开空白页
        ]
        
        print(f"[Edge] 启动远程调试模式 (端口: {self.debug_port})...")
        
        self.edge_process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        )
        
        print(f"[Edge] ✓ 已启动 (PID: {self.edge_process.pid})")
        print(f"[Edge] 等待浏览器初始化...")
        time.sleep(3)
    
    def start_flask(self):
        """启动 Flask 应用"""
        print(f"\n[Flask] 启动应用服务器...")
        
        # 使用 Python 启动 app.py
        cmd = [sys.executable, 'app.py']
        
        self.flask_process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1
        )
        
        print(f"[Flask] ✓ 已启动 (PID: {self.flask_process.pid})")
        
        # 在后台线程中打印 Flask 输出
        def print_flask_output():
            for line in self.flask_process.stdout:
                print(f"[Flask] {line.rstrip()}")
        
        Thread(target=print_flask_output, daemon=True).start()
        
        # 等待 Flask 启动
        print(f"[Flask] 等待服务器就绪...")
        time.sleep(3)
        
        # 验证服务器是否启动
        import requests
        for i in range(10):
            try:
                resp = requests.get("http://127.0.0.1:5005", timeout=1)
                if resp.status_code == 200:
                    print(f"[Flask] ✓ 服务器已就绪")
                    return True
            except:
                time.sleep(1)
        
        print(f"[Flask] ⚠ 服务器启动超时")
        return False
    
    def connect_selenium(self):
        """连接 Selenium 到远程 Edge"""
        print(f"\n[Selenium] 连接到 Edge...")
        
        from selenium import webdriver
        from selenium.webdriver.edge.service import Service
        from selenium.webdriver.edge.options import Options
        
        options = Options()
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{self.debug_port}")
        
        service = Service()
        driver = webdriver.Edge(service=service, options=options)
        
        print(f"[Selenium] ✓ 已连接")
        
        # 访问应用
        url = "http://127.0.0.1:5005"
        print(f"[Selenium] 正在打开应用: {url}")
        driver.get(url)
        time.sleep(2)
        
        print(f"[Selenium] ✓ 应用已打开")
        print(f"[Selenium]   URL: {driver.current_url}")
        print(f"[Selenium]   标题: {driver.title}")
        
        return driver
    
    def cleanup(self):
        """清理资源"""
        print(f"\n{'='*60}")
        print("正在清理资源...")
        print(f"{'='*60}")
        
        if self.flask_process:
            print("[清理] 停止 Flask 服务器...")
            try:
                self.flask_process.terminate()
                self.flask_process.wait(timeout=5)
                print("[清理] ✓ Flask 已停止")
            except:
                print("[清理] ⚠ Flask 进程清理失败")
        
        if self.edge_process:
            print("[清理] 关闭 Edge 浏览器...")
            try:
                self.edge_process.terminate()
                self.edge_process.wait(timeout=5)
                print("[清理] ✓ Edge 已关闭")
            except:
                print("[清理] ⚠ Edge 进程清理失败")
        
        print("[清理] 完成")
    
    def run(self):
        """运行完整流程"""
        print("=" * 60)
        print("网点路线优化系统启动器（远程调试模式）")
        print("=" * 60)
        
        try:
            # Step 1: 启动 Edge
            self.start_edge()
            
            # Step 2: 启动 Flask
            if not self.start_flask():
                print("\n❌ Flask 启动失败")
                return
            
            # Step 3: 连接 Selenium
            driver = self.connect_selenium()
            
            # Step 4: 保持运行
            print(f"\n{'='*60}")
            print("✓ 系统启动成功！")
            print(f"{'='*60}")
            print(f"📍 访问地址: http://127.0.0.1:5005")
            print(f"🌐 浏览器已打开并连接")
            print(f"💡 按 Ctrl+C 停止所有服务")
            print(f"{'='*60}\n")
            
            # 保持运行直到用户中断
            try:
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                print("\n\n收到停止信号...")
            
        except Exception as e:
            print(f"\n❌ 启动失败: {e}")
            import traceback
            traceback.print_exc()
        finally:
            self.cleanup()


if __name__ == "__main__":
    launcher = EdgeDebugLauncher()
    
    # 注册清理函数
    atexit.register(launcher.cleanup)
    
    # 处理 Ctrl+C
    def signal_handler(sig, frame):
        print("\n\n收到中断信号...")
        launcher.cleanup()
        sys.exit(0)
    
    signal.signal(signal.SIGINT, signal_handler)
    
    # 运行
    launcher.run()
