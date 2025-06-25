import os
import re
import math
import pandas as pd

# 计算两点之间的经纬度距离（单位：米）
def haversine(lat1, lon1, lat2, lon2):
    R = 6371000  # 地球半径（米）
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    d_phi = math.radians(lat2 - lat1)
    d_lambda = math.radians(lon2 - lon1)
    a = math.sin(d_phi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(d_lambda / 2) ** 2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

# 提取BLHA中的纬度、经度、高程
def extract_blha_and_type(content):
    blha_match = re.search(r'BLHA=([\d\.\-,]+)', content)
    entity_match = re.search(r'ENTITYNAME\s*=\s*(\w+)', content)
    if blha_match and entity_match:
        parts = blha_match.group(1).split(',')
        if len(parts) >= 3:
            lat = float(parts[0])
            lon = float(parts[1])
            alt = float(parts[2])
            return entity_match.group(1), lat, lon, alt
    return None, None, None, None

# 扫描CBM目录，提取WIRE_DEVICE 和 TOWER 信息
def scan_cbm_directory(directory):
    wire_data, tower_data = [], []
    seen_tower_files = set()

    for fname in os.listdir(directory):
        if not fname.endswith('.cbm'):
            continue
        fpath = os.path.join(directory, fname)
        with open(fpath, 'r', encoding='utf-8') as f:
            content = f.read()
            entity, lat, lon, alt = extract_blha_and_type(content)
            if entity == 'Wire_Device':
                wire_data.append((fname, lat, lon, alt))
            elif entity == 'F4System' and 'GROUPTYPE=TOWER' in content:
                if fname not in seen_tower_files:
                    tower_data.append((fname, lat, lon, alt))
                    seen_tower_files.add(fname)
    return wire_data, tower_data

# 匹配最近的TOWER并计算地面海拔
def match_wire_to_tower(wires, towers):
    results = []
    for w_file, w_lat, w_lon, w_alt in wires:
        min_dist = float('inf')
        matched = (None, None, None, None, None)
        for t_file, t_lat, t_lon, t_height in towers:
            dist = haversine(w_lat, w_lon, t_lat, t_lon)
            if dist < min_dist:
                min_dist = dist
                ground_alt = w_alt - t_height if w_alt and t_height else None
                matched = (t_file, t_lat, t_lon, t_height, ground_alt, dist)
        results.append({
            'Wire_Device': w_file,
            'Tower': matched[0],
            'Wire_Altitude': w_alt,
            'Tower_Height': matched[3],
            'Ground_Altitude': matched[4],
            'Distance(m)': round(matched[5], 3) if matched[5] is not None else None
        })
    return pd.DataFrame(results)

# 主函数入口
if __name__ == '__main__':
    cbm_dir = r'E:\tower\平江电厂\Cbm'
    wire_list, tower_list = scan_cbm_directory(cbm_dir)
    df_result = match_wire_to_tower(wire_list, tower_list)
    df_result.to_excel("matched_results_with_ground_altitude.xlsx", index=False)
    print(f"处理完成，共匹配 {len(df_result)} 条记录，结果已保存为 Excel 文件。")
