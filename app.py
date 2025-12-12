# app.py
from __future__ import annotations

import os
import math
from typing import List, Dict, Any, Tuple

import pandas as pd
import requests
from flask import Flask, request, jsonify, render_template

app = Flask(__name__)

# 从环境变量读取（推荐），也可直接写死
# 如果环境变量未设置，使用默认值（与 static/config.js 中的 webAk 保持一致）
BAIDU_WEB_AK = os.getenv("BAIDU_WEB_AK", "PnhCYT0obcdXPMchgzYz8QE4Y5ezbq36")

DIRECTIONLITE_URL = "https://api.map.baidu.com/directionlite/v1/driving"


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
    返回：polyline(list of [lng,lat]), distance(m), duration(s)
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
        "tactics": 0,
    }
    resp = requests.get(DIRECTIONLITE_URL, params=params, timeout=20)
    resp.raise_for_status()
    data = resp.json()

    if data.get("status") != 0:
        raise RuntimeError(f"百度路线规划失败：status={data.get('status')} message={data.get('message')}")

    route = data["result"]["routes"][0]
    dist = int(route.get("distance", 0))
    dur = int(route.get("duration", 0))

    poly = []
    for st in route.get("steps", []) or []:
        path = st.get("path", "")
        if not path:
            continue
        # path: "lng,lat;lng,lat;..."
        for pair in path.split(";"):
            if not pair:
                continue
            lng_s, lat_s = pair.split(",")
            poly.append([float(lng_s), float(lat_s)])

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

    return {
        "route": route,
        "polyline": polyline_all,
        "legs": legs,
        "total_distance": total_distance,
        "total_duration": total_duration,
    }


@app.get("/")
def home():
    return render_template("index.html")


@app.post("/upload_excel")
def upload_excel():
    try:
        f = request.files.get("file")
        if not f:
            return jsonify({"error": "未收到文件"}), 400

        locs = _read_excel_locations(f.stream)
        if not locs:
            return jsonify({"error": "未解析到有效网点数据（请检查经纬度、名称列）"}), 400

        return jsonify({"locations": locs})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.post("/calculate")
def calculate():
    try:
        payload = request.get_json(force=True)
        locs = payload.get("locations", [])
        if not isinstance(locs, list) or len(locs) < 2:
            return jsonify({"error": "至少需要2个网点"}), 400

        # 确保字段完整
        route = []
        for p in locs:
            route.append({
                "lng": float(p["lng"]),
                "lat": float(p["lat"]),
                "name": str(p.get("name", "")).strip(),
                "remark": str(p.get("remark", "")).strip(),
            })
        if any(not p["name"] for p in route):
            return jsonify({"error": "存在空的网点名称，请检查输入"}), 400

        result = _build_route_result(route)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.post("/optimize")
def optimize():
    try:
        payload = request.get_json(force=True)
        locs = payload.get("locations", [])
        start_name = payload.get("start_name")

        if not isinstance(locs, list) or len(locs) < 2:
            return jsonify({"error": "至少需要2个网点"}), 400

        pts = []
        for p in locs:
            pts.append({
                "lng": float(p["lng"]),
                "lat": float(p["lat"]),
                "name": str(p.get("name", "")).strip(),
                "remark": str(p.get("remark", "")).strip(),
            })

        route = _nearest_neighbor_order(pts, start_name if start_name else None)
        result = _build_route_result(route)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 400


if __name__ == "__main__":
    # 建议：export BAIDU_WEB_AK=xxxxxx
    # 或 Windows：set BAIDU_WEB_AK=xxxxxx
    app.run(host="127.0.0.1", port=5004, debug=True)