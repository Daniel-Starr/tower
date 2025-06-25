import os
import re
import math
import pandas as pd

# 设置 CBM 文件夹路径
cbm_folder = r"E:\tower\平江电厂\Cbm"

# Haversine 距离计算函数（单位：米）
def haversine(lat1, lon1, lat2, lon2):
    R = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1 - a))

# 提取 BLHA 字段信息
def parse_blha_from_content(content):
    match = re.search(r'BLHA\s*=\s*([\d\.\-]+),([\d\.\-]+),([\d\.\-]+),([\d\.\-]+)', content)
    if match:
        return tuple(map(float, match.groups()))
    return None

# 判断是否为 TOWER 类型
def is_tower(content):
    return 'GROUPTYPE=TOWER' in content

# 判断是否为 Wire_Device 类型
def is_wire_device(content):
    return 'ENTITYNAME=Wire_Device' in content

# 遍历目录下所有 .cbm 文件
tower_data, wire_data = [], []
for filename in os.listdir(cbm_folder):
    if filename.endswith(".cbm"):
        path = os.path.join(cbm_folder, filename)
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        blha = parse_blha_from_content(content)
        if not blha:
            continue
        lat, lon, alt, _ = blha
        if is_tower(content):
            tower_data.append({'filename': filename, 'lat': lat, 'lon': lon, 'height': alt})
        elif is_wire_device(content):
            wire_data.append({'filename': filename, 'lat': lat, 'lon': lon, 'altitude': alt})

# 为每个 Wire_Device 找最近的 Tower（50米内），并避免重复匹配同一 Tower 文件
matched_towers = set()
matches = []
for wire in wire_data:
    closest = None
    min_dist = float('inf')
    for tower in tower_data:
        if tower['filename'] in matched_towers:
            continue
        dist = haversine(wire['lat'], wire['lon'], tower['lat'], tower['lon'])
        if dist < min_dist and dist <= 50:
            min_dist = dist
            closest = tower
    if closest:
        matched_towers.add(closest['filename'])
        ground_altitude = wire['altitude'] - closest['height']
        matches.append({
            'Wire_Device': wire['filename'],
            'Tower': closest['filename'],
            'Wire_Altitude': wire['altitude'],
            'Tower_Height': closest['height'],
            'Ground_Altitude': round(ground_altitude, 6),
            'Distance(m)': round(min_dist, 2)
        })

# 保存为 Excel
df = pd.DataFrame(matches)
df.to_excel("calculated_ground_altitudes.xlsx", index=False)
print("✅ 匹配完成（去重），结果已保存为 calculated_ground_altitudes.xlsx")
