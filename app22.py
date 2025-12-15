# app.py
"""
网点路线优化系统 - 百度地图版
功能：网点路线规划、优化、最远网点连线显示
"""

from __future__ import annotations

import os
import sys
import subprocess
import math
import threading
import time
from typing import List, Dict, Any, Tuple, Optional

import pandas as pd
import requests
from flask import Flask, request, jsonify, render_template

# Selenium 相关导入（用于打开浏览器）
try:
    from selenium import webdriver
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.edge.options import Options
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False
    print("⚠️ 警告：未安装 selenium，将无法自动打开浏览器")
    print("   建议安装：pip install selenium")

# ==================== 配置常量 ====================
# 服务器配置
HOST = "127.0.0.1"
PORT = 5005
DEBUG_MODE = True

# 全局浏览器实例（用于截图功能复用）
_global_browser_driver = None
_global_browser_process = None  # 保存浏览器进程对象
_browser_lock = threading.Lock()

# 远程调试配置
EDGE_DEBUG_PORT = 9222  # Edge 远程调试端口

# 百度地图API配置
BAIDU_WEB_AK = os.getenv("BAIDU_WEB_AK", "PnhCYT0obcdXPMchgzYz8QE4Y5ezbq36")
DIRECTIONLITE_URL = "https://api.map.baidu.com/directionlite/v1/driving"

# API请求超时时间（秒）
API_TIMEOUT = 20

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


def _find_edge_path():
    """
    查找 Edge 浏览器路径
    
    Returns:
        Edge 浏览器路径，如果未找到返回 None
    """
    edge_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    
    for path in edge_paths:
        if os.path.exists(path):
            return path
    
    return None


def _start_edge_with_remote_debugging(debug_port=9222, user_data_dir=None):
    """
    启动带远程调试的 Edge 浏览器
    
    Args:
        debug_port: 远程调试端口
        user_data_dir: 用户数据目录（None 则使用临时目录）
    
    Returns:
        subprocess.Popen 对象，如果失败返回 None
    """
    edge_path = _find_edge_path()
    if not edge_path:
        print("[浏览器] ❌ 未找到 Edge 浏览器，请检查安装路径")
        return None
    
    # 用户数据目录（避免与正常使用的 Edge 冲突）
    if user_data_dir is None:
        user_data_dir = os.path.join(os.getenv('TEMP', os.getcwd()), 'EdgeDebugProfile_RouteSystem')
    
    # 确保目录存在
    os.makedirs(user_data_dir, exist_ok=True)
    
    # 启动命令
    cmd = [
        edge_path,
        f'--remote-debugging-port={debug_port}',
        f'--user-data-dir={user_data_dir}',
        '--no-first-run',
        '--no-default-browser-check',
        '--disable-extensions',
        '--window-size=1920,1080',
        '--disable-blink-features=AutomationControlled',
    ]
    
    try:
        print(f"[浏览器] 正在启动 Edge（远程调试模式，端口: {debug_port}）...")
        # 启动进程（Windows 下隐藏控制台窗口）
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        )
        
        print(f"[浏览器] ✓ Edge 已启动 (PID: {process.pid})")
        print(f"[浏览器] 等待浏览器初始化...")
        time.sleep(3)  # 等待浏览器完全启动
        
        return process
    except Exception as e:
        print(f"[浏览器] ❌ 启动 Edge 失败: {e}")
        return None


def _connect_to_remote_edge(debug_port=9222):
    """
    连接到远程调试模式的 Edge 浏览器
    
    Args:
        debug_port: 远程调试端口
    
    Returns:
        WebDriver 实例，如果失败返回 None
    """
    try:
        print(f"[浏览器] 正在连接到远程 Edge (端口: {debug_port})...")
        
        # 配置选项（连接到远程调试浏览器时，只需要 debuggerAddress）
        options = Options()
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
        # 注意：连接到已存在的浏览器时，不需要其他启动选项
        
        # 创建驱动（连接到已运行的浏览器）
        service = Service()
        driver = webdriver.Edge(service=service, options=options)
        driver.set_page_load_timeout(120)
        driver.implicitly_wait(5)
        
        print(f"[浏览器] ✓ 已连接到 Edge 浏览器")
        print(f"[浏览器] 当前 URL: {driver.current_url}")
        
        return driver
    except Exception as e:
        print(f"[浏览器] ❌ 连接到 Edge 失败: {e}")
        return None


def _create_browser_instance():
    """
    创建新的浏览器实例（使用远程调试模式）
    返回: webdriver实例，如果失败返回None
    """
    global _global_browser_driver, _global_browser_process
    
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
            
            # 如果已有浏览器进程，先尝试关闭
            if _global_browser_process is not None:
                try:
                    _global_browser_process.terminate()
                    _global_browser_process.wait(timeout=5)
                except:
                    try:
                        _global_browser_process.kill()
                    except:
                        pass
                _global_browser_process = None
            
            # 启动带远程调试的 Edge
            process = _start_edge_with_remote_debugging(EDGE_DEBUG_PORT)
            if process is None:
                return None
            
            _global_browser_process = process
            
            # 连接到远程调试的 Edge
            driver = _connect_to_remote_edge(EDGE_DEBUG_PORT)
            if driver is None:
                # 连接失败，关闭进程
                try:
                    process.terminate()
                except:
                    pass
                _global_browser_process = None
                return None
            
            # 访问页面
            url = f"http://{HOST}:{PORT}"
            print(f"[浏览器] 正在访问: {url}")
            driver.get(url)
            
            # 保存到全局变量
            _global_browser_driver = driver
            
            print(f"✓ 已使用 Selenium 连接到 Edge 浏览器: {url}")
            print(f"   浏览器实例已保存，截图功能将复用此实例")
            
            return driver
            
    except Exception as e:
        print(f"[浏览器] ⚠️ 创建浏览器实例失败: {e}")
        _global_browser_driver = None
        _global_browser_process = None
        return None


def _check_browser_instance():
    """
    检查浏览器实例是否有效
    如果无效，尝试重新创建
    返回: 有效的浏览器实例，如果失败返回None
    """
    global _global_browser_driver
    
    # 如果浏览器实例不存在，创建新的
    if _global_browser_driver is None:
        print("[浏览器检查] 浏览器实例不存在，正在创建...")
        return _create_browser_instance()
    
    # 检查浏览器实例是否有效
    try:
        # 尝试获取当前URL来验证浏览器是否仍然有效
        _global_browser_driver.current_url
        return _global_browser_driver
    except Exception as e:
        print(f"[浏览器检查] ⚠️ 浏览器实例无效: {e}")
        print("[浏览器检查] 正在重新创建浏览器实例...")
        return _create_browser_instance()


def _safe_float(x) -> float:
    try:
        return float(x)
    except Exception:
        return float("nan")


def _read_excel_locations(file_stream) -> List[Dict[str, Any]]:
    """
    读取Excel文件，解析网点数据
    支持列：经度、纬度、网点名称、备注(可选)、网组(可选)
    
    Returns:
        网点列表，每个网点包含：lng, lat, name, remark, group
    """
    df = pd.read_excel(file_stream)

    # 兼容列名（严格按中文列名最稳）
    # 必需：经度、纬度、网点名称
    needed = {"经度", "纬度", "网点名称"}
    cols = set(df.columns.astype(str))
    missing = needed - cols
    if missing:
        raise ValueError(f"Excel缺少列：{', '.join(missing)}。需要：经度、纬度、网点名称；备注、网组可选。")

    if "备注" not in df.columns:
        df["备注"] = ""
    if "网组" not in df.columns:
        df["网组"] = ""

    locations = []
    for _, r in df.iterrows():
        lng = _safe_float(r["经度"])
        lat = _safe_float(r["纬度"])
        name = "" if pd.isna(r["网点名称"]) else str(r["网点名称"]).strip()
        remark = "" if pd.isna(r["备注"]) else str(r["备注"]).strip()
        group = "" if pd.isna(r["网组"]) else str(r["网组"]).strip()
        if not name:
            continue
        if math.isnan(lng) or math.isnan(lat):
            continue
        locations.append({"lng": lng, "lat": lat, "name": name, "remark": remark, "group": group})
    return locations


def _call_driving_leg(a: Dict[str, Any], b: Dict[str, Any]) -> Tuple[List[List[float]], int, int]:
    """
    调用百度地图API获取两点之间的驾车路线
    
    Args:
        a: 起点，包含 lng, lat 字段
        b: 终点，包含 lng, lat 字段
    
    Returns:
        Tuple[polyline, distance, duration]:
        - polyline: 路线点列表 [[lng, lat], ...]
        - distance: 距离（米）
        - duration: 时间（秒）
    
    Raises:
        RuntimeError: API调用失败或返回错误
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
    
    try:
        resp = requests.get(DIRECTIONLITE_URL, params=params, timeout=API_TIMEOUT)
        resp.raise_for_status()
        data = resp.json()
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"百度地图API请求失败: {str(e)}")
    except ValueError as e:
        raise RuntimeError(f"百度地图API响应解析失败: {str(e)}")

    if data.get("status") != 0:
        error_msg = data.get("message", "未知错误")
        raise RuntimeError(f"百度路线规划失败：status={data.get('status')}, message={error_msg}")

    # 检查返回数据格式
    if "result" not in data or "routes" not in data["result"] or not data["result"]["routes"]:
        raise RuntimeError("百度地图API返回数据格式错误：缺少路线信息")

    route = data["result"]["routes"][0]
    dist = int(route.get("distance", 0))
    dur = int(route.get("duration", 0))

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

        # 按网组分组
        groups = {}
        for loc in locs:
            group = loc.get("group", "").strip()
            if not group:
                group = "未分组"  # 如果没有网组，归为"未分组"
            if group not in groups:
                groups[group] = []
            groups[group].append(loc)
        
        return jsonify({
            "locations": locs,
            "count": len(locs),
            "groups": groups,
            "group_count": len(groups)
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
        
        # 获取当前应用URL
        url = f"http://{HOST}:{PORT}"
        
        # 截图保存目录（打包后保存到exe所在目录的"网点图"文件夹）
        base_dir = _get_base_dir()
        save_dir = os.path.join(base_dir, "网点图")
        
        # 检查并确保浏览器实例有效（如果失效会自动重新创建）
        driver_instance = _check_browser_instance()
        
        if driver_instance is None:
            error_msg = "无法创建浏览器实例。请确保已安装selenium和Edge浏览器，并检查EdgeDriver是否正确安装。"
            print(f"[截图API] ❌ {error_msg}")
            return jsonify({"error": error_msg}), 500
        
        print("[截图API] ✓ 使用浏览器实例进行截图")
        
        # 获取网组名称（用于截图文件命名）
        group_name = data.get('group_name', '')
        
        # 执行截图（传递UI状态、浏览器实例和网组名称，等待3秒，确保页面和控制面板滚动完成）
        # 如果截图失败且是因为浏览器实例失效，会尝试重新创建并重试一次
        try:
            filepath = capture_screenshot_sync(
                url, 
                save_dir=save_dir, 
                wait_time=3, 
                ui_state=ui_state,
                driver_instance=driver_instance,
                group_name=group_name
            )
        except Exception as e:
            error_str = str(e)
            # 如果错误信息包含"invalid session id"、"no such window"或"浏览器实例无效"，尝试重新创建浏览器并重试
            if any(keyword in error_str.lower() for keyword in ["invalid session id", "no such window", "浏览器实例无效", "浏览器会话"]):
                print("[截图API] ⚠️ 检测到浏览器会话失效，正在重新创建浏览器实例并重试...")
                driver_instance = _create_browser_instance()
                if driver_instance is None:
                    return jsonify({"error": "浏览器实例失效且无法重新创建，请重新启动程序"}), 500
                # 重试截图
                try:
                    filepath = capture_screenshot_sync(
                        url, 
                        save_dir=save_dir, 
                        wait_time=3, 
                        ui_state=ui_state,
                        driver_instance=driver_instance,
                        group_name=group_name
                    )
                except Exception as retry_e:
                    # 重试也失败，返回错误
                    return jsonify({"error": f"截图失败（重试后仍失败）: {str(retry_e)}"}), 500
            else:
                # 其他错误直接抛出
                raise
        
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


def _cleanup_browser_resources():
    """清理浏览器资源"""
    global _global_browser_driver, _global_browser_process
    print("\n[清理] 正在清理浏览器资源...")
    
    if _global_browser_driver is not None:
        try:
            print("[清理] 断开 Selenium 连接...")
            _global_browser_driver.quit()
            print("[清理] ✓ Selenium 已断开")
        except:
            pass
        _global_browser_driver = None
    
    if _global_browser_process is not None:
        try:
            print("[清理] 关闭 Edge 浏览器进程...")
            _global_browser_process.terminate()
            _global_browser_process.wait(timeout=5)
            print("[清理] ✓ Edge 进程已关闭")
        except:
            try:
                _global_browser_process.kill()
                print("[清理] ✓ Edge 进程已强制关闭")
            except:
                pass
        _global_browser_process = None


if __name__ == "__main__":
    """
    主程序入口
    启动Flask服务器并自动打开浏览器
    """
    import os
    
    # 使用 Selenium 打开浏览器的函数
    def open_browser():
        """延迟打开浏览器，确保服务器已启动，使用远程调试模式启动 Edge 供截图功能复用"""
        time.sleep(3)  # 等待服务器启动
        url = f"http://{HOST}:{PORT}"
        
        if not SELENIUM_AVAILABLE:
            print(f"⚠ Selenium 未安装，无法自动打开浏览器")
            print(f"   请手动访问: {url}")
            return
        
        # 检查是否已有有效的浏览器实例
        driver = _check_browser_instance()
        if driver is None:
            print(f"⚠ 无法创建浏览器实例")
            print(f"   请手动访问: {url}")
        else:
            # 如果浏览器已打开但不在正确的URL，导航到正确页面
            try:
                current_url = driver.current_url
                if current_url != url:
                    print(f"[浏览器] 浏览器已打开，正在导航到: {url}")
                    driver.get(url)
                else:
                    print(f"✓ 浏览器已打开并位于: {url}")
            except Exception as e:
                print(f"⚠ 访问页面失败: {e}")
                # 尝试重新创建浏览器实例
                _create_browser_instance()
    
    # 打印启动信息
    print("=" * 60)
    print("🚀 网点路线优化系统正在启动...")
    print(f"📍 访问地址: http://{HOST}:{PORT}")
    print(f"🔑 API密钥: {'已配置' if BAIDU_WEB_AK else '未配置'}")
    print(f"🐛 调试模式: {'开启' if DEBUG_MODE else '关闭'}")
    print("=" * 60)
    print("💡 提示：按 Ctrl+C 停止服务器")
    print("=" * 60)
    
    try:
        # 只在主进程中打开浏览器（避免reloader导致重复打开）
        # WERKZEUG_RUN_MAIN 只在reloader子进程中为'true'
        # 主进程中没有这个环境变量，所以只在主进程中打开浏览器
        if os.environ.get('WERKZEUG_RUN_MAIN') != 'true':
            # 这是主进程，打开浏览器
            browser_thread = threading.Thread(target=open_browser)
            browser_thread.daemon = True
            browser_thread.start()
        # 如果是reloader子进程，不打开浏览器
        
        # 禁用reloader以避免重复打开浏览器，但保留debug功能
        app.run(host=HOST, port=PORT, debug=DEBUG_MODE, use_reloader=False)
    except OSError as e:
        if "Address already in use" in str(e) or "address is already in use" in str(e).lower():
            print(f"\n❌ 错误：端口 {PORT} 已被占用")
            print(f"   请关闭占用该端口的程序，或修改 PORT 配置")
        else:
            print(f"\n❌ 启动失败: {e}")
    except KeyboardInterrupt:
        print("\n\n👋 服务器已停止")
    except Exception as e:
        print(f"\n❌ 发生错误: {e}")
    finally:
        # 清理浏览器资源
        _cleanup_browser_resources()