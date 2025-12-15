# app.py
"""
网点路线优化系统 - 百度地图版
功能：网点路线规划、优化、最远网点连线显示
"""

from __future__ import annotations

import os
import sys
import math
import platform
import socket
import threading
import time
import tempfile
from typing import List, Dict, Any, Tuple, Optional

import pandas as pd
import requests
from flask import Flask, request, jsonify, render_template

# Selenium 相关导入（用于打开浏览器）
try:
    from selenium import webdriver
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.edge.options import Options
    from selenium.webdriver.common.by import By
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False
    By = None
    print("⚠️ 警告：未安装 selenium，将无法自动打开浏览器")
    print("   建议安装：pip install selenium")

# ==================== 配置常量 ====================
# 服务器配置
HOST = "127.0.0.1"
PORT = 5006
DEBUG_MODE = True

# 实际使用的端口（可能在启动时自动调整）
_actual_port = PORT

# 全局浏览器实例（用于截图功能复用）
_global_browser_driver = None
_browser_lock = threading.Lock()

# 百度地图API配置
BAIDU_WEB_AK = os.getenv("BAIDU_WEB_AK", "PnhCYT0obcdXPMchgzYz8QE4Y5ezbq36")
DIRECTIONLITE_URL = "https://api.map.baidu.com/directionlite/v1/driving"
# 行政区划查询API（使用Geocoding API的逆地理编码功能）
GEOCODING_URL = "https://api.map.baidu.com/geocoding/v3/"

# API请求超时时间（秒）
API_TIMEOUT = 20
# API请求重试次数
API_RETRY_COUNT = 3
# API请求重试延迟（秒）
API_RETRY_DELAY = 1

# ==================== Flask应用初始化 ====================
app = Flask(__name__)


def _require_ak():
    if not BAIDU_WEB_AK:
        raise RuntimeError("后端未配置 BAIDU_WEB_AK。请设置环境变量 BAIDU_WEB_AK 或在 app.py 中写入。")


def _get_base_dir():
    """
    获取程序基础目录
    在打包成exe后，返回exe所在目录；在开发环境中，返回脚本所在目录
    """
    if getattr(sys, 'frozen', False):
        # 打包成exe后，使用exe所在目录
        return os.path.dirname(sys.executable)
    else:
        # 开发环境，使用脚本所在目录
        return os.path.dirname(os.path.abspath(__file__))


def _get_actual_port():
    """
    获取实际使用的端口号
    如果端口被占用，可能会自动调整到其他可用端口
    """
    global _actual_port
    return _actual_port


def _is_port_available(host: str, port: int) -> bool:
    """
    检查端口是否可用
    
    Args:
        host: 主机地址
        port: 端口号
    
    Returns:
        如果端口可用返回 True，否则返回 False
    """
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)
            result = s.connect_ex((host, port))
            return result != 0  # 0 表示连接成功（端口被占用）
    except Exception:
        return False


def _find_available_port(host: str, start_port: int, max_attempts: int = 10) -> int:
    """
    查找可用端口
    
    Args:
        host: 主机地址
        start_port: 起始端口号
        max_attempts: 最大尝试次数
    
    Returns:
        可用的端口号，如果找不到则返回 None
    """
    for i in range(max_attempts):
        port = start_port + i
        if _is_port_available(host, port):
            return port
    return None


def _get_edge_binary_path():
    """
    根据操作系统获取 Edge 浏览器的可执行文件路径
    支持多种检测方式，提高跨电脑兼容性
    
    Returns:
        Edge 浏览器路径，如果未找到返回 None
    """
    system = platform.system()
    
    if system == "Windows":
        # Windows 系统下的 Edge 路径（按优先级排序）
        edge_paths = [
            # 标准安装路径
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            # 用户安装路径
            os.path.expanduser(r"~\AppData\Local\Microsoft\Edge\Application\msedge.exe"),
            # 可能的其他路径
            r"C:\Program Files\Microsoft\Edge Beta\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge Dev\Application\msedge.exe",
        ]
        
        # 方法1：直接检查路径
        for path in edge_paths:
            if os.path.exists(path):
                print(f"[系统检测] 检测到 Windows 系统，使用 Edge 路径: {path}")
                return path
        
        # 方法2：通过注册表查找（如果直接路径找不到）
        try:
            import winreg
            # 检查注册表中的Edge安装路径
            reg_paths = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"),
            ]
            for hkey, reg_path in reg_paths:
                try:
                    key = winreg.OpenKey(hkey, reg_path)
                    edge_path = winreg.QueryValue(key, None)
                    winreg.CloseKey(key)
                    if edge_path and os.path.exists(edge_path):
                        print(f"[系统检测] 通过注册表找到 Edge 路径: {edge_path}")
                        return edge_path
                except (FileNotFoundError, OSError):
                    continue
        except ImportError:
            pass  # winreg 在某些Python版本可能不可用
        except Exception as e:
            print(f"[系统检测] 注册表查找失败: {e}")
        
        # 方法3：通过环境变量或系统PATH查找
        try:
            import shutil
            edge_cmd = shutil.which("msedge")
            if edge_cmd and os.path.exists(edge_cmd):
                print(f"[系统检测] 通过系统PATH找到 Edge 路径: {edge_cmd}")
                return edge_cmd
        except Exception as e:
            print(f"[系统检测] PATH查找失败: {e}")
        
        print("[系统检测] ⚠️ Windows 系统下未找到 Edge 浏览器")
        print("[系统检测] 提示：请确保已安装 Microsoft Edge 浏览器")
        return None
    
    elif system == "Darwin":  # macOS
        # macOS 系统下的 Edge 路径
        edge_paths = [
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            os.path.expanduser("~/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"),
        ]
        for path in edge_paths:
            if os.path.exists(path):
                print(f"[系统检测] 检测到 macOS 系统，使用 Edge 路径: {path}")
                return path
        print("[系统检测] ⚠️ macOS 系统下未找到 Edge 浏览器")
        return None
    
    else:
        # Linux 或其他系统
        print(f"[系统检测] ⚠️ 未支持的操作系统: {system}")
        return None


def _create_browser_instance():
    """
    创建新的浏览器实例
    返回: webdriver实例，如果失败返回None
    """
    global _global_browser_driver
    
    if not SELENIUM_AVAILABLE:
        print("[浏览器] ⚠️ Selenium 未安装，无法创建浏览器实例")
        return None
    
    try:
        with _browser_lock:
            # 如果已有浏览器实例，先尝试关闭
            if _global_browser_driver is not None:
                try:
                    _global_browser_driver.quit()
                except:
                    pass
                _global_browser_driver = None
            
            print(f"[浏览器] 正在使用 Selenium 启动 Edge 浏览器...")
            
            # 获取 Edge 浏览器路径
            edge_binary_path = _get_edge_binary_path()
            
            # 配置 Edge 浏览器选项（隐藏自动化标识，防止窗口被关闭）
            edge_options = Options()
            
            # 如果找到了 Edge 路径，设置浏览器可执行文件路径
            if edge_binary_path:
                try:
                    edge_options.binary_location = edge_binary_path
                    print(f"[浏览器] 已设置 Edge 浏览器路径: {edge_binary_path}")
                except Exception as e:
                    print(f"[浏览器] ⚠️ 设置 Edge 路径失败: {e}，将使用系统默认路径")
            else:
                print("[浏览器] ⚠️ 未找到 Edge 浏览器路径，将使用系统默认路径")
                print("[浏览器] 提示：Selenium 将尝试自动查找 Edge 浏览器")
            
            edge_options.add_argument('--window-size=1920,1080')
            edge_options.add_argument('--disable-blink-features=AutomationControlled')
            edge_options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
            edge_options.add_experimental_option('useAutomationExtension', False)
            edge_options.page_load_strategy = 'normal'
            # 防止浏览器被意外关闭
            edge_options.add_argument('--disable-infobars')  # 隐藏信息栏
            edge_options.add_argument('--no-first-run')  # 跳过首次运行
            edge_options.add_argument('--no-default-browser-check')  # 跳过默认浏览器检查
            # 禁用各种提示和弹窗
            edge_options.add_argument('--disable-popup-blocking')  # 禁用弹窗阻止
            edge_options.add_argument('--disable-notifications')  # 禁用通知
            edge_options.add_argument('--disable-save-password-bubble')  # 禁用密码保存提示
            edge_options.add_argument('--disable-single-click-autofill')  # 禁用自动填充提示
            edge_options.add_argument('--disable-translate')  # 禁用翻译提示
            edge_options.add_argument('--disable-features=TranslateUI')  # 禁用翻译UI
            edge_options.add_argument('--disable-component-update')  # 禁用组件更新提示
            # 添加首选项来禁用重置设置提示
            try:
                prefs = {
                    'profile.default_content_setting_values.notifications': 2,  # 禁用通知
                    'profile.default_content_settings.popups': 0,  # 允许弹窗（避免阻止）
                    'credentials_enable_service': False,  # 禁用凭据服务
                    'profile.password_manager_enabled': False,  # 禁用密码管理器
                }
                edge_options.add_experimental_option('prefs', prefs)
            except Exception as e:
                print(f"[浏览器] ⚠️ 设置首选项失败: {e}，继续使用默认配置")
            # 保持浏览器打开（即使所有标签页关闭）
            edge_options.add_experimental_option('detach', True)  # 保持浏览器进程运行
            
            # 启动浏览器
            try:
                service = Service()
                print("[浏览器] 正在启动 Edge 浏览器...")
                driver = webdriver.Edge(service=service, options=edge_options)
                driver.set_page_load_timeout(20)
                driver.implicitly_wait(5)
                
                # 关闭所有额外的标签页（只保留第一个）
                try:
                    window_handles = driver.window_handles
                    if len(window_handles) > 1:
                        print(f"[浏览器] 检测到 {len(window_handles)} 个标签页，关闭多余的标签页...")
                        # 切换到第一个窗口
                        driver.switch_to.window(window_handles[0])
                        # 关闭其他窗口
                        for handle in window_handles[1:]:
                            try:
                                driver.switch_to.window(handle)
                                driver.close()
                            except:
                                pass
                        # 切换回第一个窗口
                        driver.switch_to.window(window_handles[0])
                        print("[浏览器] ✓ 已关闭多余的标签页")
                except Exception as e:
                    print(f"[浏览器] ⚠️ 关闭多余标签页时出错（继续）: {e}")
                
                # 访问页面
                url = f"http://{HOST}:{_get_actual_port()}"
                print(f"[浏览器] 正在访问: {url}")
                driver.get(url)
                
                # 再次检查并关闭可能新打开的标签页
                try:
                    time.sleep(0.5)  # 等待一下，确保所有标签页都已打开
                    window_handles = driver.window_handles
                    if len(window_handles) > 1:
                        print(f"[浏览器] 检测到访问后仍有 {len(window_handles)} 个标签页，关闭多余的...")
                        # 找到包含目标URL的窗口
                        target_handle = None
                        for handle in window_handles:
                            driver.switch_to.window(handle)
                            if url in driver.current_url or HOST in driver.current_url:
                                target_handle = handle
                                break
                        
                        # 切换到目标窗口，关闭其他窗口
                        if target_handle:
                            driver.switch_to.window(target_handle)
                            for handle in window_handles:
                                if handle != target_handle:
                                    try:
                                        driver.switch_to.window(handle)
                                        driver.close()
                                    except:
                                        pass
                            driver.switch_to.window(target_handle)
                        else:
                            # 如果找不到目标窗口，保留第一个，关闭其他的
                            driver.switch_to.window(window_handles[0])
                            for handle in window_handles[1:]:
                                try:
                                    driver.switch_to.window(handle)
                                    driver.close()
                                except:
                                    pass
                            driver.switch_to.window(window_handles[0])
                        print("[浏览器] ✓ 已清理多余的标签页")
                except Exception as e:
                    print(f"[浏览器] ⚠️ 清理标签页时出错（继续）: {e}")
                
                # 保存到全局变量
                _global_browser_driver = driver
                
                print(f"✓ 已使用 Selenium 打开 Edge 浏览器: {url}")
                print(f"   浏览器实例已保存，截图功能将复用此实例")
                print(f"   ⚠️ 重要提示：请勿关闭此浏览器窗口，否则截图功能将无法正常工作！")
                
                return driver
            except Exception as e:
                error_msg = str(e)
                print(f"[浏览器] ❌ 启动 Edge 浏览器失败: {error_msg}")
                
                # 提供详细的错误诊断信息
                if "WebDriver" in error_msg or "driver" in error_msg.lower():
                    print("[浏览器] 诊断：可能是 Edge WebDriver 版本不匹配")
                    print("[浏览器] 解决方案：")
                    print("   1. 确保已安装最新版本的 Microsoft Edge 浏览器")
                    print("   2. Selenium 4.x 会自动管理 WebDriver，但需要网络连接下载")
                    print("   3. 如果网络受限，请手动下载匹配的 EdgeDriver")
                
                if "path" in error_msg.lower() or "not found" in error_msg.lower():
                    print("[浏览器] 诊断：可能是 Edge 浏览器路径问题")
                    print("[浏览器] 解决方案：")
                    print("   1. 确保已安装 Microsoft Edge 浏览器")
                    print("   2. 尝试重新安装 Edge 浏览器")
                
                if "permission" in error_msg.lower() or "access" in error_msg.lower():
                    print("[浏览器] 诊断：可能是权限问题")
                    print("[浏览器] 解决方案：")
                    print("   1. 尝试以管理员身份运行程序")
                    print("   2. 检查防火墙和杀毒软件设置")
                
                raise
            
    except Exception as e:
        print(f"[浏览器] ⚠️ 创建浏览器实例失败: {e}")
        _global_browser_driver = None
        return None


def _check_browser_instance(create_if_missing: bool = True):
    """
    检查浏览器实例是否有效
    如果无效，尝试重新创建（仅在需要时）
    
    Args:
        create_if_missing: 如果浏览器实例不存在，是否创建新实例（默认True）
                          设置为False时，如果实例不存在或无效，返回None而不创建新窗口
    
    返回: 有效的浏览器实例，如果失败返回None
    """
    global _global_browser_driver
    
    # 如果浏览器实例不存在
    if _global_browser_driver is None:
        if create_if_missing:
            print("[浏览器检查] 浏览器实例不存在，正在创建...")
            return _create_browser_instance()
        else:
            # 不创建新窗口，直接返回None
            return None
    
    # 检查浏览器实例是否有效（包括窗口是否仍然打开）
    try:
        # 检查窗口句柄是否存在（如果窗口被关闭，这个会失败）
        window_handles = _global_browser_driver.window_handles
        if not window_handles:
            raise Exception("浏览器窗口已被关闭（没有活动窗口）")
        
        # 切换到第一个窗口（确保有活动窗口）
        _global_browser_driver.switch_to.window(window_handles[0])
        
        # 尝试获取当前URL来验证浏览器是否仍然有效
        current_url = _global_browser_driver.current_url
        target_url = f"http://{HOST}:{_get_actual_port()}"
        
        # 检查页面是否已经加载了应用（通过检查关键元素）
        # 如果页面已经加载，就不刷新，避免丢失已渲染的内容
        try:
            # 检查控制面板是否存在（说明应用已加载）
            if By is not None:
                control_panel = _global_browser_driver.find_elements(By.ID, "control-panel")
            else:
                control_panel = _global_browser_driver.find_elements("id", "control-panel")
            # 如果当前URL包含目标URL的基础部分，且控制面板存在，说明页面已加载
            if control_panel and (target_url in current_url or current_url.startswith(target_url.split('?')[0].split('#')[0])):
                print(f"[浏览器检查] ✓ 页面已加载应用，无需刷新（当前URL: {current_url}）")
            else:
                # 只有在页面确实不在应用页面时，才刷新
                print(f"[浏览器检查] 当前URL ({current_url}) 不是目标URL，正在导航到: {target_url}")
                _global_browser_driver.get(target_url)
                # 等待页面加载
                time.sleep(1)
        except Exception as check_e:
            # 如果检查失败，尝试导航到目标URL
            print(f"[浏览器检查] 检查页面状态失败: {check_e}，尝试导航到: {target_url}")
            _global_browser_driver.get(target_url)
            # 等待页面加载
            time.sleep(1)
        
        print(f"[浏览器检查] ✓ 浏览器实例有效，当前URL: {current_url}")
        return _global_browser_driver
    except Exception as e:
        print(f"[浏览器检查] ⚠️ 浏览器实例无效: {e}")
        # 清空无效的实例
        with _browser_lock:
            _global_browser_driver = None
        
        # 判断错误类型：如果是窗口被关闭，不创建新窗口；如果是其他错误，可以创建
        error_str = str(e).lower()
        is_window_closed = any(keyword in error_str for keyword in [
            "窗口已被关闭",
            "no such window",
            "window not found",
            "target window",
            "no window",
            "invalid session"
        ])
        
        if is_window_closed:
            # 窗口被关闭，不创建新窗口（用户可能手动关闭了）
            print("[浏览器检查] ⚠️ 浏览器窗口已被关闭，不自动创建新窗口")
            return None
        elif create_if_missing:
            # 其他错误且允许创建，尝试重新创建
            print("[浏览器检查] 正在重新创建浏览器实例...")
            return _create_browser_instance()
        else:
            # 不允许创建，返回None
            return None


def _safe_float(x) -> float:
    try:
        return float(x)
    except Exception:
        return float("nan")


def _read_excel_locations(file_stream) -> List[Dict[str, Any]]:
    """
    读取Excel文件，解析网点数据
    支持列：经度、纬度、网点名称、备注(可选)、网组(可选)、工号(可选)、姓名(可选)、县区(可选)、调整(可选)
    
    Returns:
        网点列表，每个网点包含：lng, lat, name, remark, group, employee_id, employee_name, district, adjustment
    """
    df = pd.read_excel(file_stream)

    # 兼容列名（严格按中文列名最稳）
    # 必需：经度、纬度、网点名称
    needed = {"经度", "纬度", "网点名称"}
    cols = set(df.columns.astype(str))
    missing = needed - cols
    if missing:
        raise ValueError(f"Excel缺少列：{', '.join(missing)}。需要：经度、纬度、网点名称；备注、网组、工号、姓名、县区、调整可选。")

    if "备注" not in df.columns:
        df["备注"] = ""
    if "网组" not in df.columns:
        df["网组"] = ""
    if "工号" not in df.columns:
        df["工号"] = ""
    if "姓名" not in df.columns:
        df["姓名"] = ""
    if "县区" not in df.columns:
        df["县区"] = ""
    if "调整" not in df.columns:
        df["调整"] = ""

    locations = []
    for _, r in df.iterrows():
        lng = _safe_float(r["经度"])
        lat = _safe_float(r["纬度"])
        name = "" if pd.isna(r["网点名称"]) else str(r["网点名称"]).strip()
        remark = "" if pd.isna(r["备注"]) else str(r["备注"]).strip()
        group = "" if pd.isna(r["网组"]) else str(r["网组"]).strip()
        employee_id = "" if pd.isna(r["工号"]) else str(r["工号"]).strip()
        employee_name = "" if pd.isna(r["姓名"]) else str(r["姓名"]).strip()
        district = "" if pd.isna(r["县区"]) else str(r["县区"]).strip()
        adjustment = "" if pd.isna(r["调整"]) else str(r["调整"]).strip()
        if not name:
            continue
        if math.isnan(lng) or math.isnan(lat):
            continue
        locations.append({
            "lng": lng, 
            "lat": lat, 
            "name": name, 
            "remark": remark, 
            "group": group,
            "employee_id": employee_id,
            "employee_name": employee_name,
            "district": district,
            "adjustment": adjustment
        })
    return locations


def _call_driving_leg(a: Dict[str, Any], b: Dict[str, Any]) -> Tuple[List[List[float]], int, int]:
    """
    调用百度地图API获取两点之间的驾车路线（带重试机制）
    
    Args:
        a: 起点，包含 lng, lat 字段
        b: 终点，包含 lng, lat 字段
    
    Returns:
        Tuple[polyline, distance, duration]:
        - polyline: 路线点列表 [[lng, lat], ...]
        - distance: 距离（米）
        - duration: 时间（秒）
    
    Raises:
        RuntimeError: API调用失败或返回错误（重试后仍失败）
    """
    _require_ak()

    # 注意：百度接口参数为 lat,lng
    params = {
        "ak": BAIDU_WEB_AK,
        "origin": f'{a["lat"]},{a["lng"]}',
        "destination": f'{b["lat"]},{b["lng"]}',
        "coord_type": "bd09ll",
        "ret_coordtype": "bd09ll",
        "steps_info": 1,
        "tactics": 0,  # 0=不走高速，1=最短时间，2=最短距离
    }
    
    last_exception = None
    
    # 重试机制
    for attempt in range(API_RETRY_COUNT):
        try:
            resp = requests.get(DIRECTIONLITE_URL, params=params, timeout=API_TIMEOUT)
            resp.raise_for_status()
            data = resp.json()
            
            # 检查API返回状态
            if data.get("status") != 0:
                error_msg = data.get("message", "未知错误")
                status_code = data.get("status")
                
                # 某些错误不应该重试（如参数错误、权限错误等）
                if status_code in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30]:
                    # 这些是业务逻辑错误，不需要重试
                    raise RuntimeError(f"百度路线规划失败：status={status_code}, message={error_msg}")
                
                # 其他错误（如服务器错误、限流等）可以重试
                last_exception = RuntimeError(f"百度路线规划失败：status={status_code}, message={error_msg}")
                if attempt < API_RETRY_COUNT - 1:
                    print(f"[API重试] 第 {attempt + 1} 次请求失败（status={status_code}），{API_RETRY_DELAY}秒后重试...")
                    time.sleep(API_RETRY_DELAY)
                    continue
                else:
                    raise last_exception
            
            # 检查返回数据格式
            if "result" not in data or "routes" not in data["result"] or not data["result"]["routes"]:
                last_exception = RuntimeError("百度地图API返回数据格式错误：缺少路线信息")
                if attempt < API_RETRY_COUNT - 1:
                    print(f"[API重试] 第 {attempt + 1} 次请求返回数据格式错误，{API_RETRY_DELAY}秒后重试...")
                    time.sleep(API_RETRY_DELAY)
                    continue
                else:
                    raise last_exception
            
            # 请求成功，解析数据
            route = data["result"]["routes"][0]
            dist = int(route.get("distance", 0))
            dur = int(route.get("duration", 0))
            
            # 如果之前有重试，打印成功信息
            if attempt > 0:
                print(f"[API重试] 第 {attempt + 1} 次请求成功")
            
            # 解析路线点
            poly = []
            for st in route.get("steps", []) or []:
                path = st.get("path", "")
                if not path:
                    continue
                # path格式: "lng,lat;lng,lat;..."
                for pair in path.split(";"):
                    if not pair or "," not in pair:
                        continue
                    try:
                        lng_s, lat_s = pair.split(",", 1)
                        poly.append([float(lng_s), float(lat_s)])
                    except ValueError:
                        continue  # 跳过无效的坐标点
            
            return poly, dist, dur
            
        except requests.exceptions.Timeout as e:
            last_exception = RuntimeError(f"百度地图API请求超时: {str(e)}")
            if attempt < API_RETRY_COUNT - 1:
                print(f"[API重试] 第 {attempt + 1} 次请求超时，{API_RETRY_DELAY}秒后重试...")
                time.sleep(API_RETRY_DELAY)
                continue
            else:
                raise last_exception
                
        except requests.exceptions.ConnectionError as e:
            last_exception = RuntimeError(f"百度地图API连接失败: {str(e)}")
            if attempt < API_RETRY_COUNT - 1:
                print(f"[API重试] 第 {attempt + 1} 次请求连接失败，{API_RETRY_DELAY}秒后重试...")
                time.sleep(API_RETRY_DELAY)
                continue
            else:
                raise last_exception
                
        except requests.exceptions.RequestException as e:
            last_exception = RuntimeError(f"百度地图API请求失败: {str(e)}")
            if attempt < API_RETRY_COUNT - 1:
                print(f"[API重试] 第 {attempt + 1} 次请求失败，{API_RETRY_DELAY}秒后重试...")
                time.sleep(API_RETRY_DELAY)
                continue
            else:
                raise last_exception
                
        except ValueError as e:
            last_exception = RuntimeError(f"百度地图API响应解析失败: {str(e)}")
            if attempt < API_RETRY_COUNT - 1:
                print(f"[API重试] 第 {attempt + 1} 次响应解析失败，{API_RETRY_DELAY}秒后重试...")
                time.sleep(API_RETRY_DELAY)
                continue
            else:
                raise last_exception
    
    # 所有重试都失败
    if last_exception:
        raise last_exception
    else:
        raise RuntimeError("百度地图API请求失败：未知错误")

    # 解析路线点
    poly = []
    for st in route.get("steps", []) or []:
        path = st.get("path", "")
        if not path:
            continue
        # path格式: "lng,lat;lng,lat;..."
        for pair in path.split(";"):
            if not pair or "," not in pair:
                continue
            try:
                lng_s, lat_s = pair.split(",", 1)
                poly.append([float(lng_s), float(lat_s)])
            except ValueError:
                continue  # 跳过无效的坐标点

    return poly, dist, dur


def _format_distance_m(m: int) -> str:
    if m >= 1000:
        return f"{m/1000:.2f} 公里"
    return f"{m} 米"


def _format_duration_s(s: int) -> str:
    h = s // 3600
    mm = (s % 3600) // 60
    if h > 0:
        return f"{h}小时{mm}分钟"
    return f"{mm}分钟"


def _calculate_straight_distance(loc1: Dict[str, Any], loc2: Dict[str, Any]) -> float:
    """
    计算两个网点之间的直线距离（米）
    使用Haversine公式计算球面距离
    """
    # 地球半径（米）
    R = 6371000
    
    lat1 = math.radians(loc1["lat"])
    lat2 = math.radians(loc2["lat"])
    delta_lat = math.radians(loc2["lat"] - loc1["lat"])
    delta_lng = math.radians(loc2["lng"] - loc1["lng"])
    
    a = math.sin(delta_lat / 2) ** 2 + \
        math.cos(lat1) * math.cos(lat2) * math.sin(delta_lng / 2) ** 2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    
    return R * c


def _find_farthest_points(locs: List[Dict[str, Any]]) -> Optional[Tuple[Dict[str, Any], Dict[str, Any], float]]:
    """
    找到两个最远的网点
    返回：(点1, 点2, 直线距离(米))
    """
    if len(locs) < 2:
        return None
    
    max_dist = 0
    farthest_pair = None
    
    for i in range(len(locs)):
        for j in range(i + 1, len(locs)):
            dist = _calculate_straight_distance(locs[i], locs[j])
            if dist > max_dist:
                max_dist = dist
                farthest_pair = (locs[i], locs[j], dist)
    
    return farthest_pair


def _nearest_neighbor_order(locs: List[Dict[str, Any]], start_name: str | None) -> List[Dict[str, Any]]:
    """
    简单最近邻：用于“优化路线”的顺序建议（不是严格TSP最优，但够实用且很快）
    """
    if len(locs) <= 2:
        return locs[:]

    remaining = locs[:]

    # 选择起点
    start_idx = 0
    if start_name:
        for i, p in enumerate(remaining):
            if p["name"] == start_name:
                start_idx = i
                break

    route = [remaining.pop(start_idx)]

    def dist2(p, q):
        dx = p["lng"] - q["lng"]
        dy = p["lat"] - q["lat"]
        return dx*dx + dy*dy

    while remaining:
        last = route[-1]
        best_i = min(range(len(remaining)), key=lambda i: dist2(last, remaining[i]))
        route.append(remaining.pop(best_i))

    return route


def _build_route_result(route: List[Dict[str, Any]]) -> Dict[str, Any]:
    polyline_all: List[List[float]] = []
    legs = []
    total_distance = 0
    total_duration = 0
    leg_polylines = []  # 保存每个路段的polyline，用于计算中点

    for i in range(len(route) - 1):
        a, b = route[i], route[i + 1]
        poly, dist, dur = _call_driving_leg(a, b)
        if polyline_all and poly:
            # 去重拼接点
            if polyline_all[-1] == poly[0]:
                poly = poly[1:]
        polyline_all.extend(poly)
        leg_polylines.append(poly)

        # 计算当前路段的中点坐标
        mid_point = None
        if poly and len(poly) > 0:
            mid_idx = len(poly) // 2
            mid_point = poly[mid_idx]

        legs.append({
            "from": a["name"],
            "to": b["name"],
            "distance": dist,
            "duration": dur,
            "distance_text": _format_distance_m(dist),
            "duration_text": _format_duration_s(dur),
            "mid_point": mid_point,  # 添加中点坐标用于标注距离
        })
        total_distance += dist
        total_duration += dur

    # 计算最远的两个网点
    farthest_info = None
    farthest_pair = _find_farthest_points(route)
    if farthest_pair:
        point1, point2, straight_dist = farthest_pair
        farthest_info = {
            "point1": point1,
            "point2": point2,
            "straight_distance": int(straight_dist),
            "straight_distance_text": _format_distance_m(int(straight_dist)),
        }

    return {
        "route": route,
        "polyline": polyline_all,
        "legs": legs,
        "total_distance": total_distance,
        "total_duration": total_duration,
        "farthest_points": farthest_info,
    }


@app.get("/")
def home():
    return render_template("index.html")


@app.get("/config-custom.js")
def get_custom_config():
    """
    提供config-custom.js配置文件
    优先从exe同目录或脚本目录读取config-custom.js，如果不存在则返回404
    """
    base_dir = _get_base_dir()
    config_path = os.path.join(base_dir, "config-custom.js")
    
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                content = f.read()
            from flask import Response
            return Response(content, mimetype='application/javascript')
        except Exception as e:
            print(f"[配置] 读取config-custom.js失败: {e}")
            return jsonify({"error": "读取配置文件失败"}), 500
    else:
        # 如果config-custom.js不存在，返回404，前端会fallback到static/config.js
        from flask import abort
        abort(404)


@app.post("/upload_excel")
def upload_excel():
    """
    上传Excel文件并解析网点数据
    
    Returns:
        JSON响应，包含locations列表或error信息
    """
    try:
        f = request.files.get("file")
        if not f:
            return jsonify({"error": "未收到文件"}), 400

        # 检查文件扩展名
        filename = f.filename or ""
        if not (filename.endswith('.xlsx') or filename.endswith('.xls')):
            return jsonify({"error": "文件格式错误，请上传 .xlsx 或 .xls 文件"}), 400

        locs = _read_excel_locations(f.stream)
        if not locs:
            return jsonify({"error": "未解析到有效网点数据（请检查经纬度、名称列）"}), 400

        # 首先按"调整"字段分组，然后在同字段下再按工号、网组分组
        # 结构：adjustments -> employee_id -> groups -> locations
        adjustments = {}
        for loc in locs:
            adjustment = loc.get("adjustment", "").strip()
            if not adjustment:
                adjustment = "未分类"  # 如果没有调整字段，归为"未分类"
            
            if adjustment not in adjustments:
                adjustments[adjustment] = {}
            
            employee_id = loc.get("employee_id", "").strip()
            if not employee_id:
                employee_id = "未分组"  # 如果没有工号，归为"未分组"
            
            if employee_id not in adjustments[adjustment]:
                adjustments[adjustment][employee_id] = {
                    "employee_id": employee_id,
                    "employee_name": loc.get("employee_name", "").strip(),
                    "groups": {}
                }
            
            group = loc.get("group", "").strip()
            if not group:
                group = "未分组"  # 如果没有网组，归为"未分组"
            
            if group not in adjustments[adjustment][employee_id]["groups"]:
                adjustments[adjustment][employee_id]["groups"][group] = []
            
            adjustments[adjustment][employee_id]["groups"][group].append(loc)
        
        # 保持向后兼容：按网组分组（用于旧版功能）
        groups = {}
        for loc in locs:
            group = loc.get("group", "").strip()
            if not group:
                group = "未分组"
            if group not in groups:
                groups[group] = []
            groups[group].append(loc)
        
        # 保持向后兼容：按工号分组（用于旧版功能）
        employees = {}
        for loc in locs:
            employee_id = loc.get("employee_id", "").strip()
            if employee_id:
                if employee_id not in employees:
                    employees[employee_id] = {
                        "employee_id": employee_id,
                        "employee_name": loc.get("employee_name", "").strip(),
                        "groups": {}
                    }
                group = loc.get("group", "").strip()
                if not group:
                    group = "未分组"
                if group not in employees[employee_id]["groups"]:
                    employees[employee_id]["groups"][group] = []
                employees[employee_id]["groups"][group].append(loc)
        
        return jsonify({
            "locations": locs,
            "count": len(locs),
            "groups": groups,
            "group_count": len(groups),
            "employees": employees,
            "employee_count": len(employees),
            "adjustments": adjustments,  # 新增：按调整字段分组的数据
            "adjustment_count": len(adjustments)
        })
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"文件处理失败: {str(e)}"}), 500


@app.post("/calculate")
def calculate():
    """
    按顺序计算路线（不优化顺序）
    
    Returns:
        JSON响应，包含路线结果或error信息
    """
    try:
        payload = request.get_json(force=True)
        if not payload:
            return jsonify({"error": "请求体为空"}), 400

        locs = payload.get("locations", [])
        if not isinstance(locs, list):
            return jsonify({"error": "locations必须是数组"}), 400
        
        if len(locs) < 2:
            return jsonify({"error": "至少需要2个网点"}), 400

        # 验证并格式化网点数据
        route = []
        for idx, p in enumerate(locs):
            try:
                route.append({
                    "lng": float(p["lng"]),
                    "lat": float(p["lat"]),
                    "name": str(p.get("name", "")).strip(),
                    "remark": str(p.get("remark", "")).strip(),
                })
            except (KeyError, ValueError, TypeError) as e:
                return jsonify({"error": f"第{idx+1}个网点数据格式错误: {str(e)}"}), 400

        # 验证网点名称
        if any(not p["name"] for p in route):
            return jsonify({"error": "存在空的网点名称，请检查输入"}), 400

        # 计算路线
        result = _build_route_result(route)
        
        # 调试输出
        if result.get("farthest_points"):
            fp = result["farthest_points"]
            print(f"[calculate] 最远网点: {fp['point1']['name']} <-> {fp['point2']['name']}, "
                  f"距离: {fp['straight_distance_text']}")
        
        return jsonify(result)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        return jsonify({"error": f"计算失败: {str(e)}"}), 500


@app.post("/optimize")
def optimize():
    """
    优化路线顺序（使用最近邻算法）
    
    Returns:
        JSON响应，包含优化后的路线结果或error信息
    """
    try:
        payload = request.get_json(force=True)
        if not payload:
            return jsonify({"error": "请求体为空"}), 400

        locs = payload.get("locations", [])
        start_name = payload.get("start_name")

        if not isinstance(locs, list):
            return jsonify({"error": "locations必须是数组"}), 400
        
        if len(locs) < 2:
            return jsonify({"error": "至少需要2个网点"}), 400

        # 验证并格式化网点数据
        pts = []
        for idx, p in enumerate(locs):
            try:
                pts.append({
                    "lng": float(p["lng"]),
                    "lat": float(p["lat"]),
                    "name": str(p.get("name", "")).strip(),
                    "remark": str(p.get("remark", "")).strip(),
                })
            except (KeyError, ValueError, TypeError) as e:
                return jsonify({"error": f"第{idx+1}个网点数据格式错误: {str(e)}"}), 400

        # 优化路线顺序
        route = _nearest_neighbor_order(pts, start_name if start_name else None)
        
        # 计算路线
        result = _build_route_result(route)
        
        # 调试输出
        if result.get("farthest_points"):
            fp = result["farthest_points"]
            print(f"[optimize] 最远网点: {fp['point1']['name']} <-> {fp['point2']['name']}, "
                  f"距离: {fp['straight_distance_text']}")
        
        return jsonify(result)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        return jsonify({"error": f"优化失败: {str(e)}"}), 500


@app.post("/capture_screenshot")
def capture_screenshot_endpoint():
    """
    截图API端点：使用Selenium + Edge浏览器截取当前浏览器页面viewport
    支持同步当前浏览器中的UI状态（如复选框状态）
    
    Returns:
        JSON响应，包含截图文件路径或error信息
    """
    try:
        from jietu import capture_screenshot_sync
        
        # 获取请求中的UI状态
        data = request.get_json() or {}
        ui_state = data.get('ui_state', {})
        
        # 获取网组名称、工号、姓名、调整（用于截图文件命名和路径）
        group_name = data.get('group_name', '')
        employee_id = data.get('employee_id', '')
        employee_name = data.get('employee_name', '')
        adjustment = data.get('adjustment', '')
        
        # 获取当前应用URL
        url = f"http://{HOST}:{_get_actual_port()}"
        
        # 截图保存目录（打包后保存到exe所在目录的"网组网点路线图"文件夹）
        base_dir = _get_base_dir()
        base_save_dir = os.path.join(base_dir, "网组网点路线图")
        
        # 如果有工号和姓名，创建子文件夹
        if employee_id and employee_id.strip() and employee_name and employee_name.strip():
            safe_employee_id = "".join(c for c in employee_id.strip() if c.isalnum() or c in ('-', '_', ' ')).strip().replace(' ', '_')
            safe_employee_name = "".join(c for c in employee_name.strip() if c.isalnum() or c in ('-', '_', ' ')).strip().replace(' ', '_')
            save_dir = os.path.join(base_save_dir, f"{safe_employee_id}-{safe_employee_name}")
            
            # 如果有调整字段，在工号-姓名目录下再创建调整字段子目录
            if adjustment and adjustment.strip():
                safe_adjustment = "".join(c for c in adjustment.strip() if c.isalnum() or c in ('-', '_', ' ')).strip().replace(' ', '_')
                save_dir = os.path.join(save_dir, safe_adjustment)
        else:
            save_dir = base_save_dir
        
        # 检查浏览器实例
        # 如果浏览器实例不存在，尝试创建（正常启动场景）
        # 如果浏览器实例存在但窗口被关闭，不创建新窗口（用户手动关闭的场景）
        driver_instance = _check_browser_instance(create_if_missing=True)
        
        if driver_instance is None:
            error_msg = "无法创建或访问浏览器实例。请确保：\n1. app.py已启动\n2. Edge浏览器和EdgeDriver已正确安装\n3. 如果浏览器窗口被关闭，请重新启动app.py"
            print(f"[截图API] ❌ {error_msg}")
            return jsonify({"error": error_msg}), 500
        
        # 再次验证浏览器窗口是否打开（双重检查）
        try:
            window_handles = driver_instance.window_handles
            if not window_handles:
                error_msg = "浏览器窗口已关闭。请确保浏览器窗口保持打开状态，或重新启动app.py。"
                print(f"[截图API] ❌ {error_msg}")
                return jsonify({"error": error_msg}), 500
            else:
                # 确保切换到活动窗口
                driver_instance.switch_to.window(window_handles[0])
                # 检查页面是否已经加载了应用（通过检查关键元素）
                # 如果页面已经加载，就不刷新，避免丢失已渲染的内容
                try:
                    # 检查控制面板是否存在（说明应用已加载）
                    if By is not None:
                        control_panel = driver_instance.find_elements(By.ID, "control-panel")
                    else:
                        control_panel = driver_instance.find_elements("id", "control-panel")
                    current_url = driver_instance.current_url
                    # 如果当前URL包含目标URL的基础部分，且控制面板存在，说明页面已加载
                    if control_panel and (url in current_url or current_url.startswith(f"http://{HOST}:{_get_actual_port()}")):
                        print(f"[截图API] ✓ 页面已加载，无需刷新（当前URL: {current_url}）")
                    else:
                        # 只有在页面确实不在应用页面时，才刷新
                        print(f"[截图API] 页面未加载应用，正在导航到: {url}")
                        driver_instance.get(url)
                        time.sleep(1)
                except Exception as check_e:
                    # 如果检查失败，尝试导航到目标URL
                    print(f"[截图API] 检查页面状态失败: {check_e}，尝试导航到: {url}")
                    driver_instance.get(url)
                    time.sleep(1)
        except Exception as e:
            error_msg = f"检查浏览器窗口时出错: {str(e)}。请确保浏览器窗口保持打开状态。"
            print(f"[截图API] ❌ {error_msg}")
            return jsonify({"error": error_msg}), 500
        
        print("[截图API] ✓ 使用浏览器实例进行截图")
        
        # 执行截图（传递UI状态、浏览器实例和网组名称，等待3秒，确保页面和控制面板滚动完成）
        # 如果截图失败，会尝试多次重试（最多3次）
        max_retries = 3
        retry_delay = 2  # 重试延迟（秒）
        filepath = None
        last_error = None
        
        for retry_count in range(max_retries):
            try:
                filepath = capture_screenshot_sync(
                    url,
                    save_dir=save_dir,
                    wait_time=3,
                    ui_state=ui_state,
                    driver_instance=driver_instance,
                    group_name=group_name,
                    employee_id=employee_id,
                    employee_name=employee_name,
                    adjustment=adjustment
                )
                # 截图成功，跳出重试循环
                if retry_count > 0:
                    print(f"[截图API] ✓ 第 {retry_count + 1} 次尝试成功")
                break
            except Exception as e:
                last_error = e
                error_str = str(e).lower()
                
                # 如果是最后一次重试，直接抛出错误
                if retry_count == max_retries - 1:
                    print(f"[截图API] ❌ 截图失败（已重试 {max_retries} 次）: {str(e)}")
                    break
                
                # 不重新创建浏览器，只重试（使用已打开的浏览器）
                print(f"[截图API] ⚠️ 第 {retry_count + 1} 次尝试失败: {str(e)}")
                print(f"[截图API] {retry_delay}秒后重试...")
                
                # 等待后重试
                time.sleep(retry_delay)
        
        # 如果所有重试都失败
        if filepath is None:
            return jsonify({"error": f"截图失败（已重试 {max_retries} 次）: {str(last_error)}"}), 500
        
        # 返回相对路径（基于base_dir）
        rel_path = os.path.relpath(filepath, base_dir)
        
        return jsonify({
            "success": True,
            "filepath": rel_path,
            "filename": os.path.basename(filepath),
            "message": "截图保存成功"
        })
    except ImportError as e:
        return jsonify({"error": f"截图模块导入失败: {str(e)}，请确保已安装selenium: pip install selenium。同时需要安装Edge浏览器和EdgeDriver"}), 500
    except Exception as e:
        return jsonify({"error": f"截图失败: {str(e)}"}), 500


@app.post("/get_district_boundary")
def get_district_boundary():
    """
    获取行政区划边界
    使用百度地图逆地理编码API获取坐标点所在的行政区，然后返回行政区名称
    
    Args:
        JSON请求体包含：locations (坐标点列表)
    
    Returns:
        JSON响应，包含每个坐标点对应的行政区信息
    """
    try:
        payload = request.get_json(force=True)
        if not payload:
            return jsonify({"error": "请求体为空"}), 400
        
        locations = payload.get("locations", [])
        if not isinstance(locations, list):
            return jsonify({"error": "locations必须是数组"}), 400
        
        _require_ak()
        
        district_info = []
        
        for loc in locations:
            lng = loc.get("lng")
            lat = loc.get("lat")
            
            if lng is None or lat is None:
                continue
            
            # 调用逆地理编码API获取行政区信息
            params = {
                "ak": BAIDU_WEB_AK,
                "location": f"{lat},{lng}",  # 注意：百度API参数为 lat,lng
                "output": "json",
                "coordtype": "bd09ll",
                "extensions_poi": 0,
                "extensions_road": 0,
                "extensions_town": 1,  # 返回乡镇信息
            }
            
            try:
                resp = requests.get(GEOCODING_URL, params=params, timeout=API_TIMEOUT)
                resp.raise_for_status()
                data = resp.json()
                
                if data.get("status") == 0 and "result" in data:
                    address_component = data["result"].get("addressComponent", {})
                    # 提取行政区信息：省+市+区县
                    province = address_component.get("province", "")
                    city = address_component.get("city", "")
                    district = address_component.get("district", "")
                    
                    # 组合成完整行政区名称（如：江苏南京市建邺区）
                    if province and city and district:
                        district_name = f"{province}{city}{district}"
                    elif province and district:
                        district_name = f"{province}{district}"
                    elif district:
                        district_name = district
                    else:
                        district_name = "未知区域"
                    
                    district_info.append({
                        "lng": lng,
                        "lat": lat,
                        "district": district_name,
                        "province": province,
                        "city": city,
                        "district_level": district
                    })
                else:
                    district_info.append({
                        "lng": lng,
                        "lat": lat,
                        "district": "未知区域",
                        "province": "",
                        "city": "",
                        "district_level": ""
                    })
            except Exception as e:
                print(f"[行政区查询] 查询失败 ({lng}, {lat}): {e}")
                district_info.append({
                    "lng": lng,
                    "lat": lat,
                    "district": "查询失败",
                    "province": "",
                    "city": "",
                    "district_level": ""
                })
        
        return jsonify({
            "success": True,
            "districts": district_info
        })
    except Exception as e:
        return jsonify({"error": f"获取行政区信息失败: {str(e)}"}), 500


if __name__ == "__main__":
    """
    主程序入口
    启动Flask服务器并自动打开浏览器
    """
    import os
    import traceback
    
    try:
        # 打印系统信息（用于诊断）
        print("=" * 60)
        print("系统信息:")
        print(f"  操作系统: {platform.system()} {platform.release()}")
        if hasattr(sys, 'frozen'):
            print(f"  运行模式: 打包后的EXE")
            print(f"  程序路径: {sys.executable}")
        else:
            print(f"  运行模式: Python脚本")
            print(f"  Python版本: {sys.version.split()[0]}")
            print(f"  程序路径: {os.path.abspath(__file__)}")
        print("=" * 60)
        print()
        
        # 检查端口是否可用，如果被占用则自动查找可用端口
        actual_port = PORT
        if not _is_port_available(HOST, PORT):
            print(f"⚠️ 端口 {PORT} 已被占用，正在查找可用端口...")
            available_port = _find_available_port(HOST, PORT, max_attempts=10)
            if available_port:
                actual_port = available_port
                print(f"✓ 找到可用端口: {actual_port}")
            else:
                print(f"❌ 错误：无法找到可用端口（已尝试 {PORT} 到 {PORT + 9}）")
                print(f"   请关闭占用端口的程序，或修改 PORT 配置")
                input("按回车键退出...")
                sys.exit(1)
        
        # 更新全局实际端口变量
        _actual_port = actual_port
        
        # 使用 Selenium 打开浏览器的函数（使用闭包捕获 actual_port）
        def open_browser():
            """延迟打开浏览器，确保服务器已启动，使用 Selenium 打开浏览器供截图功能复用"""
            time.sleep(1.5)  # 等待服务器启动
            url = f"http://{HOST}:{actual_port}"
            
            if not SELENIUM_AVAILABLE:
                print(f"⚠️ Selenium 未安装，无法自动打开浏览器")
                print(f"   请手动访问: {url}")
                return
            
            # 检查是否已有有效的浏览器实例
            try:
                driver = _check_browser_instance()
                if driver is None:
                    print(f"⚠️ 无法创建浏览器实例")
                    print(f"   请手动访问: {url}")
                else:
                    # 关闭多余的标签页，只保留目标页面
                    try:
                        window_handles = driver.window_handles
                        if len(window_handles) > 1:
                            # 找到包含目标URL的窗口
                            target_handle = None
                            for handle in window_handles:
                                driver.switch_to.window(handle)
                                current_url = driver.current_url
                                if url in current_url or HOST in current_url:
                                    target_handle = handle
                                    break
                            
                            # 切换到目标窗口，关闭其他窗口
                            if target_handle:
                                driver.switch_to.window(target_handle)
                                for handle in window_handles:
                                    if handle != target_handle:
                                        try:
                                            driver.switch_to.window(handle)
                                            # 检查是否是data:页面或空白页
                                            if 'data:' in driver.current_url or driver.current_url == 'about:blank' or not driver.current_url.startswith('http'):
                                                driver.close()
                                        except:
                                            pass
                                driver.switch_to.window(target_handle)
                            else:
                                # 如果找不到目标窗口，保留第一个，关闭其他的
                                driver.switch_to.window(window_handles[0])
                                for handle in window_handles[1:]:
                                    try:
                                        driver.switch_to.window(handle)
                                        if 'data:' in driver.current_url or driver.current_url == 'about:blank' or not driver.current_url.startswith('http'):
                                            driver.close()
                                    except:
                                        pass
                                driver.switch_to.window(window_handles[0])
                    except Exception as e:
                        print(f"[浏览器] ⚠️ 清理标签页时出错（继续）: {e}")
                    
                    # 如果浏览器已打开但不在正确的URL，导航到正确页面
                    try:
                        driver.switch_to.window(driver.window_handles[0])
                        current_url = driver.current_url
                        if current_url != url and not url in current_url:
                            print(f"[浏览器] 浏览器已打开，正在导航到: {url}")
                            driver.get(url)
                        else:
                            print(f"✓ 浏览器已打开并位于: {url}")
                    except Exception as e:
                        print(f"⚠️ 访问页面失败: {e}")
                        # 尝试重新创建浏览器实例
                        _create_browser_instance()
            except Exception as e:
                print(f"⚠️ 打开浏览器时出错: {e}")
                print(f"   请手动访问: {url}")
                # 只在调试模式下打印详细错误
                if DEBUG_MODE:
                    print(f"   错误详情:\n{traceback.format_exc()}")
        
        # 打印启动信息
        print("=" * 60)
        print("🚀 网点路线优化系统正在启动...")
        print(f"📍 访问地址: http://{HOST}:{actual_port}")
        print(f"🔑 API密钥: {'已配置' if BAIDU_WEB_AK else '未配置'}")
        print(f"🐛 调试模式: {'开启' if DEBUG_MODE else '关闭'}")
        print("=" * 60)
        print("💡 提示：按 Ctrl+C 停止服务器")
        print("=" * 60)
        
        # 在后台线程中打开浏览器
        browser_thread = threading.Thread(target=open_browser, daemon=True)
        browser_thread.start()
        
        # 启动Flask服务器
        try:
            app.run(host=HOST, port=actual_port, debug=DEBUG_MODE, use_reloader=False)
        except OSError as e:
            if "Address already in use" in str(e) or "address is already in use" in str(e).lower():
                print(f"\n❌ 错误：端口 {actual_port} 已被占用")
                print(f"   请关闭占用该端口的程序，或修改 PORT 配置")
            else:
                print(f"\n❌ 启动失败: {e}")
                if DEBUG_MODE:
                    print(f"   错误详情:\n{traceback.format_exc()}")
            input("按回车键退出...")
        except KeyboardInterrupt:
            print("\n\n👋 服务器已停止")
        except Exception as e:
            print(f"\n❌ 发生错误: {e}")
            if DEBUG_MODE:
                print(f"   错误详情:\n{traceback.format_exc()}")
            print("\n如果问题持续，请检查：")
            print("  1. 是否已安装 Microsoft Edge 浏览器")
            print("  2. 防火墙是否阻止了程序运行")
            print("  3. 是否有足够的系统权限")
            print("  4. 查看错误详情（如果调试模式已开启）")
            input("按回车键退出...")
    
    except Exception as e:
        print(f"\n❌ 程序启动失败: {e}")
        print(f"   错误类型: {type(e).__name__}")
        if DEBUG_MODE:
            print(f"   错误详情:\n{traceback.format_exc()}")
        print("\n常见问题解决方案：")
        print("  1. 确保已安装 Microsoft Edge 浏览器")
        print("  2. 确保已安装所有必要的系统库（Visual C++ Redistributable）")
        print("  3. 尝试以管理员身份运行")
        print("  4. 检查杀毒软件是否阻止了程序运行")
        input("按回车键退出...")
        sys.exit(1)