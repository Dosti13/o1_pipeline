"""
Join two Excel files on 'material_code'.
From file_2, only bring in 'need_review' and 'category_method' columns.
Everything else in the result comes from file_1 (unchanged).
"""

import pandas as pd

# ---- CONFIG: update these paths / sheet names as needed ----
FILE_1 = r"C:\Users\HP\Downloads\dim_material_master_qnie_csv_RESOLVED_v3_updated.xlsx"
SHEET_1 = "standardized_master"

FILE_2 = r"C:\Users\HP\Downloads\dim_material_master_qnie_csv_RESOLVED_v2 (1).xlsx"   # <-- update with actual file name
SHEET_2 = 0  # or a sheet name string, e.g. "Sheet1"

JOIN_KEY = "material_code"
COLUMNS_TO_BRING = ["needs_review", "category_method"]

OUTPUT_FILE = r"C:\Users\HP\Downloads\dim_material_master_qnie_csv_RESOLVED_v3_updated_final.xlsx"
# --------------------------------------------------------------

# Load both files
df1 = pd.read_excel(FILE_1, sheet_name=SHEET_1)
df2 = pd.read_excel(FILE_2, sheet_name=SHEET_2)

# Make sure join key has the same dtype on both sides (avoids silent join failures)
df1[JOIN_KEY] = df1[JOIN_KEY].astype(str).str.strip()
df2[JOIN_KEY] = df2[JOIN_KEY].astype(str).str.strip()

# Keep only join key + the columns we want to pull from file_2
df2_subset = df2[[JOIN_KEY] + COLUMNS_TO_BRING]

# Left join: keep all rows/columns from file_1, add only the selected columns from file_2
merged = df1.merge(df2_subset, on=JOIN_KEY, how="left")

# Quick match-rate check
matched = merged[COLUMNS_TO_BRING[0]].notna().sum()
print(f"Total rows in file_1: {len(df1)}")
print(f"Rows matched from file_2: {matched}")
print(f"Rows unmatched: {len(df1) - matched}")

merged.to_excel(OUTPUT_FILE, index=False)
print(f"Saved: {OUTPUT_FILE}")