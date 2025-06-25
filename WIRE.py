import os
import pandas as pd

def extract_wire_blha_only(cbm_folder: str, output_excel: str = "wire_blha_only.xlsx") -> None:
    result = []

    for root, _, files in os.walk(cbm_folder):
        for file in files:
            if not file.lower().endswith(".cbm"):
                continue

            file_path = os.path.join(root, file)
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = [line.strip() for line in f.readlines()]

                wire_info = {"文件名": file}
                is_wire = any("ENTITYNAME=WIRE" in line.upper() for line in lines)
                if not is_wire:
                    continue

                for line in lines:
                    upper_line = line.upper()
                    if upper_line.startswith("POINT0.BLHA="):
                        values = [float(x.strip()) for x in line.split("=")[1].split(",")]
                        if len(values) == 4:
                            wire_info["P0_纬度"] = round(values[0], 8)
                            wire_info["P0_经度"] = round(values[1], 8)
                            wire_info["P0_高程"] = round(values[2], 3)
                            wire_info["P0_偏角"] = round(values[3], 6)
                    elif upper_line.startswith("POINT1.BLHA="):
                        values = [float(x.strip()) for x in line.split("=")[1].split(",")]
                        if len(values) == 4:
                            wire_info["P1_纬度"] = round(values[0], 8)
                            wire_info["P1_经度"] = round(values[1], 8)
                            wire_info["P1_高程"] = round(values[2], 3)
                            wire_info["P1_偏角"] = round(values[3], 6)

                if "P0_纬度" in wire_info and "P1_纬度" in wire_info:
                    result.append(wire_info)

            except Exception as e:
                print(f"❌ 错误读取文件 {file_path} -> {e}")

    if result:
        df = pd.DataFrame(result)
        df.to_excel(output_excel, index=False)
        print(f"✅ 成功提取 {len(result)} 条导线端点 BLHA 信息，导出为 {output_excel}")
    else:
        print("⚠️ 没有找到任何包含导线端点 BLHA 的 WIRE 文件")

# 示例调用
if __name__ == "__main__":
    extract_wire_blha_only(r"E:\tower\平江电厂\Cbm")
