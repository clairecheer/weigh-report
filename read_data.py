import pandas as pd
import os

data_dir = r"c:\CC\Macro\data"
files = os.listdir(data_dir)
for f in sorted(files):
    path = os.path.join(data_dir, f)
    try:
        xl = pd.ExcelFile(path)
        print(f"=== {f} ===")
        print(f"  Sheets: {xl.sheet_names}")
        for sheet in xl.sheet_names[:3]:
            df = pd.read_excel(path, sheet_name=sheet, header=None)
            print(f"  Sheet [{sheet}]: shape={df.shape}")
            print(df.head(10).to_string())
            print("  ...")
        print()
    except Exception as e:
        print(f"  ERROR: {e}")
