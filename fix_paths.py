"""
运行方法：
  把此脚本放到你的 streamlit_mvtec_experiment 文件夹里，
  然后在该文件夹打开 CMD，执行：
      python fix_paths.py
会生成一个新的 Excel 文件，原文件不会被修改。
"""

import re
from pathlib import Path
import pandas as pd

WORKBOOK = Path("05_metadata/AI_质检实验题库_实验一实验二_多目标采购版.xlsx")
OUTPUT   = Path("05_metadata/AI_质检实验题库_实验一实验二_多目标采购版_fixed.xlsx")

# 需要处理"图片源路径"列的所有 sheet
SHEETS_WITH_IMAGE_PATH = ["题库总表", "Exp2_多目标采购", "Practice"]

PREFIX_PATTERN = re.compile(
    r".*?00_raw[\\/]",   # 贪婪匹配，去掉直到 00_raw\ 的所有内容
    re.IGNORECASE
)

def fix_path(value):
    if not isinstance(value, str) or value.strip() == "":
        return value
    # 去掉 00_raw\ 之前（含）的部分，再把反斜杠换成正斜杠
    cleaned = PREFIX_PATTERN.sub("", value.strip())
    cleaned = cleaned.replace("\\", "/")
    return cleaned

print(f"读取题库：{WORKBOOK}")
xl = pd.ExcelFile(WORKBOOK)
all_sheets = xl.sheet_names
print(f"共 {len(all_sheets)} 个 sheet：{all_sheets}")

writer = pd.ExcelWriter(OUTPUT, engine="openpyxl")

for sheet in all_sheets:
    df = pd.read_excel(WORKBOOK, sheet_name=sheet, header=None)
    if sheet in SHEETS_WITH_IMAGE_PATH:
        # 找到表头行（含"图片源路径"的行）
        header_row = None
        col_idx = None
        for i in range(min(10, len(df))):
            for j, val in enumerate(df.iloc[i]):
                if isinstance(val, str) and "图片源路径" in val:
                    header_row = i
                    col_idx = j
                    break
            if header_row is not None:
                break

        if col_idx is not None:
            print(f"  [{sheet}] 找到'图片源路径'列（第{col_idx+1}列，表头在第{header_row+1}行），开始替换...")
            fixed = 0
            for row_i in range(header_row + 1, len(df)):
                old = df.iat[row_i, col_idx]
                new = fix_path(old)
                if new != old:
                    df.iat[row_i, col_idx] = new
                    fixed += 1
            print(f"  [{sheet}] 替换了 {fixed} 个路径")
        else:
            print(f"  [{sheet}] 未找到'图片源路径'列，跳过")
    else:
        print(f"  [{sheet}] 无需处理，原样保留")

    df.to_excel(writer, sheet_name=sheet, index=False, header=False)

writer.close()
print(f"\n✅ 完成！新文件已保存到：{OUTPUT}")
print("请检查新文件路径是否正确，确认后可替换原文件。")
