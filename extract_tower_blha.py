import os
import pandas as pd

def extract_tower_blha(cbm_folder: str, output_excel: str = "tower_blha.xlsx") -> None:
    """
    扫描指定路径下所有 .cbm 文件，提取 GROUPTYPE=TOWER 的文件中的 BLHA 信息，
    并导出为 Excel 表格。
    """
    result = []

    for root, _, files in os.walk(cbm_folder):
        for file in files:
            if not file.lower().endswith(".cbm"):
                continue

            file_path = os.path.join(root, file)

            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = [line.strip() for line in f.readlines()]

                # 判断是否为 TOWER 文件
                is_tower = any("GROUPTYPE=TOWER" in line.upper() for line in lines)
                if not is_tower:
                    continue

                # 提取 BLHA 信息
                for line in lines:
                    if line.upper().startswith("BLHA="):
                        try:
                            parts = [float(x.strip()) for x in line.split("=")[1].split(",")]
                            if len(parts) == 4:
                                result.append({
                                    "文件名": file,
                                    "纬度": round(parts[0], 8),
                                    "经度": round(parts[1], 8),
                                    "高程（m）": round(parts[2], 3),
                                    "北方向偏角（°）": round(parts[3], 6),
                                })
                        except Exception as e:
                            print(f"⚠️ 解析 BLHA 出错: {file_path} -> {e}")
                        break  # 每个文件只提取一个 BLHA

            except Exception as e:
                print(f"❌ 无法读取文件: {file_path} -> {e}")

    # 输出为 Excel
    if result:
        df = pd.DataFrame(result)
        df.to_excel(output_excel, index=False)
        print(f"✅ 成功提取 {len(result)} 条塔的 BLHA 信息，已导出到 {output_excel}")
    else:
        print("⚠️ 没有找到任何包含 BLHA 的 TOWER 文件")

# 示例调用（请修改为你本地的路径）
if __name__ == "__main__":
    extract_tower_blha(r"E:\tower\平江电厂\Cbm")
