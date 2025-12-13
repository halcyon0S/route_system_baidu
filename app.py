# app.py
"""
网点路线优化系统 - 百度地图版
功能：网点路线规划、优化、最远网点连线显示
"""

from __future__ import annotations

import os
import math
import webbrowser
import threading
import time
from typing import List, Dict, Any, Tuple, Optional

import pandas as pd
import requests
from flask import Flask, request, jsonify, render_template

# ==================== 配置常量 ====================
# 服务器配置
HOST = "127.0.0.1"
PORT = 5005
DEBUG_MODE = True

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


def _safe_float(x) -> float:
    try:
        return float(x)
    except Exception:
        return float("nan")


def _read_excel_locations(file_stream) -> List[Dict[str, Any]]:
    df = pd.read_excel(file_stream)

    # 兼容列名（严格按中文列名最稳）
    # 必需：经度、纬度、网点名称
    needed = {"经度", "纬度", "网点名称"}
    cols = set(df.columns.astype(str))
    missing = needed - cols
    if missing:
        raise ValueError(f"Excel缺少列：{', '.join(missing)}。需要：经度、纬度、网点名称；备注可选。")

    if "备注" not in df.columns:
        df["备注"] = ""

    locations = []
    for _, r in df.iterrows():
        lng = _safe_float(r["经度"])
        lat = _safe_float(r["纬度"])
        name = "" if pd.isna(r["网点名称"]) else str(r["网点名称"]).strip()
        remark = "" if pd.isna(r["备注"]) else str(r["备注"]).strip()
        if not name:
            continue
        if math.isnan(lng) or math.isnan(lat):
            continue
        locations.append({"lng": lng, "lat": lat, "name": name, "remark": remark})
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

        return jsonify({"locations": locs, "count": len(locs)})
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


if __name__ == "__main__":
    """
    主程序入口
    启动Flask服务器并自动打开浏览器
    """
    import os
    
    # 自动打开浏览器的函数
    def open_browser():
        """延迟打开浏览器，确保服务器已启动"""
        time.sleep(1.5)  # 等待服务器启动
        url = f"http://{HOST}:{PORT}"
        try:
            webbrowser.open(url)
            print(f"✓ 已自动打开浏览器: {url}")
        except Exception as e:
            print(f"⚠ 自动打开浏览器失败: {e}")
            print(f"   请手动访问: {url}")
    
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