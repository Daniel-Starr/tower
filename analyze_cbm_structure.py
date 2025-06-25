import os
import re
import pandas as pd
from pathlib import Path

# 设置 CBM 文件目录
cbm_dir = Path("E:/tower/平江电厂/Cbm")

# 映射关系：Wire_Device 文件名 → 所属 Tower 文件名 和 Tower 高度
wire_to_tower_map = {}
tower_height_map = {}

# 第一步：扫描所有 TOWER.cbm 文件，记录 SUBDEVICE 和 TOWER 高度
for file in cbm_dir.glob("*.cbm"):
    try:
        with open(file, encoding="utf-8", errors="ignore") as f:
            content = f.read()
    except:
        continue

    if "GROUPTYPE=TOWER" in content:
        # 提取塔高
        tower_blha_match = re.search(r"BLHA=([^\n\r]+)", content)
        tower_height = None
        if tower_blha_match:
            try:
                tower_blha = [float(x.strip()) for x in tower_blha_match.group(1).split(",")]
                tower_height = tower_blha[2]
            except:
                continue
        tower_name = file.name
        tower_height_map[tower_name] = tower_height

        # 提取 SUBDEVICE 映射
        sub_matches = re.findall(r"SUBDEVICE\d+=(.*?\.cbm)", content)
        for sub in sub_matches:
            wire_to_tower_map[sub.strip()] = (tower_name, tower_height)

# 第二步：处理所有 Wire_Device 文件，提取高程并计算地面标高
results = []
for file in cbm_dir.glob("*.cbm"):
    try:
        with open(file, encoding="utf-8", errors="ignore") as f:
            content = f.read()
    except:
        continue

    if "ENTITYNAME=Wire_Device" in content:
        blha_match = re.search(r"BLHA=([^\n\r]+)", content)
        if blha_match:
            try:
                wire_blha = [float(x.strip()) for x in blha_match.group(1).split(",")]
                wire_elevation = wire_blha[2]
                wire_name = file.name

                # 查找对应 TOWER
                if wire_name in wire_to_tower_map:
                    tower_name, tower_height = wire_to_tower_map[wire_name]
                    if tower_height is not None:
                        ground_alt = wire_elevation - tower_height
                        results.append({
                            "WIRE_DEVICE": wire_name,
                            "TOWER": tower_name,
                            "Wire_Altitude": wire_elevation,
                            "Tower_Height": tower_height,
                            "Ground_Altitude": ground_alt
                        })
            except:
                continue

# 保存结果
df = pd.DataFrame(results)
output_path = "calculated_ground_altitudes.xlsx"
df.to_excel(output_path, index=False)

print(f"✅ 提取完成，结果保存在：{output_path}")
