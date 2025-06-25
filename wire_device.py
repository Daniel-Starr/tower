import os
import pandas as pd
from typing import List, Dict

def extract_wire_device_from_cbm(cbm_folder: str, output_excel: str = "wire_device_components.xlsx") -> None:
    device_data_list: List[Dict] = []

    def parse_cbm_file(filepath: str):
        try:
            with open(filepath, 'r', encoding='utf-8') as file:
                lines = file.readlines()

            i = 0
            while i < len(lines):
                line = lines[i].strip()
                if line.lower().startswith("entityname="):
                    entity_name = line.split("=")[1].strip()
                    if entity_name.strip().upper() == "WIRE_DEVICE":
                        component = {
                            "ENTITYNAME": "WIRE_DEVICE",
                            "CBM路径": filepath,
                            "原始字段": {}
                        }
                        i += 1
                        while i < len(lines) and not lines[i].lower().startswith("entityname="):
                            current_line = lines[i].strip()
                            if "=" in current_line:
                                key, value = current_line.split("=", 1)
                                key = key.strip()
                                value = value.strip()

                                if key.lower() == "blha":
                                    try:
                                        lat, lng, elev, angle = [x.strip() for x in value.split(",")]
                                        component["纬度"] = f"{float(lat):.8f}"
                                        component["经度"] = f"{float(lng):.8f}"
                                        component["高程"] = f"{float(elev):.3f}"
                                        component["北方向偏角"] = f"{float(angle):.6f}"
                                    except Exception as e:
                                        print(f"⚠️ BLHA 解析失败: {value} in {filepath} - {e}")
                                        component["纬度"] = component["经度"] = component["高程"] = component["北方向偏角"] = "解析错误"
                                else:
                                    component[key] = value
                                    component["原始字段"][key] = value
                            i += 1
                        device_data_list.append(component)
                        continue
                i += 1
        except Exception as e:
            print(f"Failed to parse {filepath}: {e}")

    for root, _, files in os.walk(cbm_folder):
        for file in files:
            if file.lower().endswith(".cbm"):
                full_path = os.path.join(root, file)
                parse_cbm_file(full_path)

    if device_data_list:
        # 展开所有字段
        df = pd.DataFrame(device_data_list)
        df.to_excel(output_excel, index=False)
        print(f"✅ WIRE_DEVICE 导出完成: {output_excel}")
    else:
        print("⚠️ 未发现符合条件的 WIRE_DEVICE 部件。")

if __name__ == "__main__":
    extract_wire_device_from_cbm(
        cbm_folder=r"E:\tower\平江电厂\Cbm",
        output_excel="wire_device_components.xlsx"
    )
