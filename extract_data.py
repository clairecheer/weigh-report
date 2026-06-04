# -*- coding: utf-8 -*-
import pandas as pd
import os
import json

data_dir = r"c:\CC\Macro\data"

def safe_read(path, sheet_name=0):
    """Read Excel and return dataframe with proper encoding"""
    df = pd.read_excel(path, sheet_name=sheet_name, header=None)
    return df

results = {}

# 1. PMI
print("=" * 60)
print("PMI - 采购经理指数")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "PMI.xlsx"))
# Find header row and data
for i in range(min(10, len(df))):
    row = df.iloc[i]
    print(f"Row {i}: {[str(x)[:50] for x in row[:5]]}")

# Print last 6 data rows
print("\nLast 6 months PMI data:")
for i in range(max(1, len(df)-7), len(df)):
    row = df.iloc[i]
    print(f"  {row[0]} | 制造业PMI: {row[1]} | 生产: {row[2]} | 新订单: {row[3]} | 出厂价格: {row[8]} | 原材料价格: {row[9]}")

print("\n")

# 2. PPI - focus on sectors relevant to weighing
print("=" * 60)
print("PPI - 工业生产者出厂价格指数 (Key sectors)")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "PPI.xlsx"))
# Header row 0 has all categories
header = df.iloc[0]
# Find relevant columns: 黑色金属 (ferrous metals), 有色金属 (non-ferrous), 非金属矿物 (non-metallic minerals),
# 通用设备 (general equipment), 专用设备 (special equipment), 汽车 (automotive), 金属制品 (metal products)
# 化学原料 (chemical raw materials), 橡胶和塑料 (rubber & plastics)
# These are columns related to industries that use industrial weighing
target_keywords = ['黑色金属', '有色金属', '非金属矿物', '通用设备', '专用设备', '汽车制造', '金属制品',
                   '化学原料', '橡胶', '煤炭', '石油', '非金属矿', '废弃资源', '总指数']

relevant_cols = [0]  # Period
for j in range(1, len(header)):
    h = str(header[j])
    for kw in target_keywords:
        if kw in h:
            relevant_cols.append(j)
            break

print(f"Found {len(relevant_cols)} relevant PPI columns")
# Print last 6 rows for key sectors
for i in range(max(1, len(df)-7), len(df)):
    row = df.iloc[i]
    period = row[0]
    vals = []
    for j in relevant_cols[:12]:
        vals.append(f"{str(header[j])[:30]}={row[j]}")
    print(f"  {period}: {' | '.join(vals)}")

print("\n")

# 3. Steel Price
print("=" * 60)
print("Steel Price - 钢铁价格")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "Steel Price.xlsx"))
# Find data rows (skip header rows)
print("Last 10 data rows:")
data_start = 0
for i in range(len(df)):
    val = df.iloc[i, 0]
    if isinstance(val, str) and '202' in val:
        data_start = i
        break

for i in range(max(data_start, len(df)-10), len(df)):
    row = df.iloc[i]
    print(f"  {row[0]} | 价格:{row[1]} | 涨跌:{row[3]} | 月环比:{row[7]}% | 同比:{row[9]}%")

print("\n")

# 4. Fixed Asset Investment
print("=" * 60)
print("Fixed Asset Investment - Key sectors")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "固定资产投资.xlsx"))
header = df.iloc[0]
# Key columns: total, second industry, manufacturing, equipment purchase,
# mining, ferrous metals, non-ferrous metals, non-metallic minerals
# general equipment, special equipment, metal products, transport/storage
target_kws = ['固定资产投资额', '第二产业', '制造业', '设备工器具', '采矿业',
              '黑色金属', '有色金属', '非金属矿物', '通用设备', '专用设备', '金属制品',
              '交通运输', '仓储', '煤炭', '化学原料', '橡胶和塑料', '电气机械',
              '电力', '水利', '房地产']

relevant_cols = [0]
for j in range(1, len(header)):
    h = str(header[j])
    for kw in target_kws:
        if kw in h:
            relevant_cols.append(j)
            break

print(f"Found {len(relevant_cols)} relevant FAI columns")
for i in range(max(1, len(df)-7), len(df)):
    row = df.iloc[i]
    if pd.notna(row[0]):
        print(f"  {row[0]}: 总FAI={row[1]}% | 第二产业={row[4]}% | 制造业={row[15]}% | 设备购置={row[12]}% | 交运仓储={row[67]}%")

print("\n")

# 5. Real Estate
print("=" * 60)
print("Real Estate - 房地产 (Key indicators)")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "房地产.xlsx"))
for i in range(min(12, len(df))):
    row = df.iloc[i]
    print(f"Row {i}: {[str(x)[:60] for x in row[:6]]}")

print("\nLast 6 rows key data:")
for i in range(max(1, len(df)-7), len(df)):
    row = df.iloc[i]
    vals = [str(row[j])[:30] for j in range(min(12, len(row))) if pd.notna(row[j])]
    print(f"  {' | '.join(vals)}")

print("\n")

# 6. Industrial Electricity
print("=" * 60)
print("Industrial Electricity - 工业用电")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "工业用电.xlsx"))
for i in range(min(15, len(df))):
    row = df.iloc[i]
    print(f"Row {i}: {[str(x)[:60] for x in row]}")
print(f"... Total rows: {len(df)}")

print("\n")

# 7. Industrial Medium-Long Term Loans
print("=" * 60)
print("Industrial Loans - 工业中长期贷款")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "工业中长期贷款.xlsx"))
for i in range(len(df)):
    row = df.iloc[i]
    print(f"Row {i}: {[str(x)[:60] for x in row]}")

print("\n")

# 8. Industrial Products Output - key sheets
print("=" * 60)
print("Industrial Products Output - 工业主要产品产量")
print("=" * 60)
xl = pd.ExcelFile(os.path.join(data_dir, "工业主要产品产量.xlsx"))
print(f"Sheets: {xl.sheet_names}")
# Read first sheet (产量/production)
for sheet in xl.sheet_names[:3]:
    df = safe_read(os.path.join(data_dir, "工业主要产品产量.xlsx"), sheet)
    print(f"\nSheet [{sheet}]: shape={df.shape}")
    for i in range(min(8, len(df))):
        row = df.iloc[i]
        print(f"  Row {i}: {[str(x)[:50] for x in row[:4]]}")

print("\n")

# 9. Import/Export
print("=" * 60)
print("Import/Export - 货物进出口总额")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "货物进出口总额.xlsx"))
for i in range(min(10, len(df))):
    row = df.iloc[i]
    print(f"Row {i}: {[str(x)[:50] for x in row[:6]]}")
print("\nLast 6 rows:")
for i in range(max(1, len(df)-7), len(df)):
    row = df.iloc[i]
    vals = [str(row[j])[:30] for j in range(len(row)) if pd.notna(row[j])]
    print(f"  {' | '.join(vals[:8])}")

print("\n")

# 10. Retail Sales
print("=" * 60)
print("Retail Sales - 社会消费品零售总额")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "社会消费品零售总额.xlsx"))
for i in range(min(8, len(df))):
    row = df.iloc[i]
    print(f"Row {i}: {[str(x)[:50] for x in row[:6]]}")

print("\n")

# 11. Logistics
print("=" * 60)
print("Logistics - 物流")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "物流.xlsx"))
for i in range(min(10, len(df))):
    row = df.iloc[i]
    print(f"Row {i}: {[str(x)[:60] for x in row[:6]]}")
print(f"... Total rows: {len(df)}")

print("\n")

# 12. 线索机遇数量 (Lead/Opportunity Count)
print("=" * 60)
print("Leads/Opportunities - 线索机遇数量")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "线索机遇数量.xlsx"))
for i in range(min(10, len(df))):
    row = df.iloc[i]
    print(f"Row {i}: {[str(x)[:60] for x in row[:8]]}")
print(f"... Total rows: {len(df)}")

print("\n")

# 13. Industry Value Added
print("=" * 60)
print("Industry Value Added - 行业增加值")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "行业增加值.xlsx"))
for i in range(min(10, len(df))):
    row = df.iloc[i]
    print(f"Row {i}: {[str(x)[:60] for x in row[:6]]}")
print(f"... Total rows: {len(df)}")

print("\n")

# 14. CPI
print("=" * 60)
print("CPI - 居民消费价格指数")
print("=" * 60)
df = safe_read(os.path.join(data_dir, "CPI.xlsx"))
print("Last 6 months CPI:")
for i in range(max(1, len(df)-7), len(df)):
    row = df.iloc[i]
    print(f"  {row[0]} | CPI: {row[1]}% | 食品: {row[2]}% | 交通通信: {row[6]}%")

print("\n=== ALL DATA EXTRACTED ===")
