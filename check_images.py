"""
在你的 streamlit_mvtec_experiment 文件夹中运行：
    python check_images.py
会列出所有在 00_raw 中找不到的图片。
"""
import pandas as pd
from pathlib import Path

df = pd.read_excel('05_metadata/MVTec_实验题库_完整版_解释优化版.xlsx', sheet_name='题库总表')
missing = []
found = []
for _, row in df.iterrows():
    p = str(row['图片源路径']).replace('\\', '/')
    if '00_raw/' in p:
        rel = p.split('00_raw/', 1)[1]
        candidate = Path('00_raw') / rel
        if not candidate.exists():
            missing.append((row['题号'], str(candidate)))
        else:
            found.append(row['题号'])

print(f"✅ 找到图片：{len(found)} 张")
print(f"❌ 缺失图片：{len(missing)} 张")
for q, path in missing:
    print(f"   {q}: {path}")
