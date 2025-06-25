import os
import pandas as pd

def extract_wire_device_blha_precise(cbm_folder: str, output_excel: str = "wire_device_blha.xlsx") -> None:
    results = []

    for root, _, files in os.walk(cbm_folder):
        for file in files:
            if not file.lower().endswith(".cbm"):
                continue

            file_path = os.path.join(root, file)
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = [line.strip() for line in f.readlines()]

                i = 0
                while i < len(lines):
                    if lines[i].upper().startswith("BLHA="):
                        # 先取 BLHA 值
                        try:
                            blha_line = lines[i]
                            parts = [float(x.strip()) for x in blha_line.split("=")[1].split(",")]
                            if len(parts) != 4:
                                i += 1
                                continue
                        except:
                            i += 1
                            continue

                        # 向后寻找 ENTITYNAME=Wire_Device
                        for j in range(i+1, min(i+5, len(lines))):
                            if lines[j].upper().startswith("ENTITYNAME=WIRE_DEVICE"):
                                results.append({
                                    "文件名": file,
                                    "纬度": round(parts[0], 8),
                                    "经度": round(parts[1], 8),
                                    "高程（m）": round(parts[2], 3),
                                    "北方向偏角（°）": round(parts[3], 6),
                                })
                                break
                    i += 1

            except Exception as e:
                print(f"❌ 错误读取文件 {file_path}: {e}")

    if results:
        df = pd.DataFrame(results)
        df.to_excel(output_excel, index=False)
        print(f"✅ 已成功提取 {len(results)} 条 BLHA 信息，保存为 {output_excel}")
    else:
        print("⚠️ 未找到任何 Wire_Device 的 BLHA 信息")

# 示例调用
if __name__ == "__main__":
    extract_wire_device_blha_precise(r"E:\tower\平江电厂\Cbm")
