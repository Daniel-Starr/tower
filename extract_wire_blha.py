import os
import pandas as pd

def extract_wire_info_with_blha(cbm_folder: str, output_excel: str = "wire_info.xlsx") -> None:
    result = []

    for root, _, files in os.walk(cbm_folder):
        for file in files:
            if not file.lower().endswith(".cbm"):
                continue

            file_path = os.path.join(root, file)
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = [line.strip() for line in f.readlines()]

                current_wire = {}
                is_wire = False

                for line in lines:
                    upper_line = line.upper()

                    if upper_line.startswith("ENTITYNAME=WIRE"):
                        is_wire = True
                        current_wire["文件名"] = file

                    elif is_wire:
                        if upper_line.startswith("BASEFAMILY="):
                            current_wire["BASEFAMILY"] = line.split("=")[1].strip()
                        elif upper_line.startswith("OBJECTMODELPOINTER="):
                            current_wire["OBJECTMODELPOINTER"] = line.split("=")[1].strip()
                        elif upper_line.startswith("KVALUE="):
                            current_wire["KVALUE"] = line.split("=")[1].strip()
                        elif upper_line.startswith("SPLIT="):
                            current_wire["SPLIT"] = line.split("=")[1].strip()
                        elif upper_line.startswith("POINT0.BLHA="):
                            blha = line.split("=")[1].strip()
                            current_wire["POINT0.BLHA"] = blha
                            parts = [float(x.strip()) for x in blha.split(",")]
                            if len(parts) == 4:
                                current_wire["P0_纬度"] = round(parts[0], 8)
                                current_wire["P0_经度"] = round(parts[1], 8)
                                current_wire["P0_高程"] = round(parts[2], 3)
                                current_wire["P0_偏角"] = round(parts[3], 6)
                        elif upper_line.startswith("POINT1.BLHA="):
                            blha = line.split("=")[1].strip()
                            current_wire["POINT1.BLHA"] = blha
                            parts = [float(x.strip()) for x in blha.split(",")]
                            if len(parts) == 4:
                                current_wire["P1_纬度"] = round(parts[0], 8)
                                current_wire["P1_经度"] = round(parts[1], 8)
                                current_wire["P1_高程"] = round(parts[2], 3)
                                current_wire["P1_偏角"] = round(parts[3], 6)
                        elif upper_line.startswith("ENTITYNAME=") and not upper_line.startswith("ENTITYNAME=WIRE"):
                            if current_wire:
                                result.append(current_wire)
                                current_wire = {}
                                is_wire = False

                if is_wire and current_wire:
                    result.append(current_wire)

            except Exception as e:
                print(f"❌ 错误读取文件 {file_path} -> {e}")

    if result:
        df = pd.DataFrame(result)
        df.to_excel(output_excel, index=False)
        print(f"✅ 提取完成，共 {len(result)} 条记录，已导出至 {output_excel}")
    else:
        print("⚠️ 未找到任何包含 WIRE 的 .cbm 文件")

# 示例调用（替换为你实际的文件夹路径）
if __name__ == "__main__":
    extract_wire_info_with_blha(r"E:\tower\平江电厂\Cbm")
