import pandas as pd
import os

# ==============================
# FILE PATHS
# ==============================

MAIN_FILE = r"C:\Users\HP\Desktop\functionand sp.csv"

FILES = [
    r"C:\Users\HP\Downloads\Al Meera July 2024.xlsx",
    r"C:\Users\HP\Downloads\Talabat March.xlsx",
    r"C:\Users\HP\Downloads\fwqniemtselloutdashboardsession2\C4 sep 24.xlsx",
    r"C:\Users\HP\Downloads\fwqniemtselloutdashboardsession2\Grand Mall APR 2024.xlsx",
    r"C:\Users\HP\Downloads\fwqniemtselloutdashboardsession2\QNIE..xlsx",
    r"C:\Users\HP\Downloads\fwqniemtselloutdashboardsession2\RawSparData.xlsx"
]

OUTPUT_FILE = r"C:\Users\HP\Desktop\final_output.csv"


# ==============================
# SAFE FILE READER
# ==============================

def read_file(file):
    ext = os.path.splitext(file)[1].lower()

    try:
        if ext == ".csv":
            try:
                return pd.read_csv(file, dtype=str, encoding="utf-8")
            except:
                return pd.read_csv(file, dtype=str, encoding="latin-1")

        elif ext in [".xlsx", ".xls"]:
            return pd.read_excel(file, dtype=str, engine="openpyxl")

        return pd.DataFrame()

    except Exception as e:
        print(f"Error reading {file}: {e}")
        return pd.DataFrame()


# ==============================
# LOAD DATA
# ==============================

print("Loading files...")

df_main = read_file(MAIN_FILE)
df_list = [read_file(f) for f in FILES]

df_all = pd.concat(df_list, ignore_index=True)

print("Files loaded successfully!")


# ==============================
# NORMALIZE COLUMNS
# ==============================

def normalize(df):
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )
    return df


df_main = normalize(df_main)
df_all = normalize(df_all)


# ==============================
# SAFE CLEAN FUNCTION (FIXED)
# ==============================

def clean_df(df):
    for col in df.columns:
        df[col] = df[col].apply(
            lambda x: x.strip().lower() if isinstance(x, str) else x
        )
    return df


df_main = clean_df(df_main)
df_all = clean_df(df_all)


# ==============================
# AUTO COLUMN DETECTION
# ==============================

def find_col(df, keywords):
    for col in df.columns:
        for k in keywords:
            if k in col:
                return col
    return None


# MAIN FILE KEYS
m_barcode = find_col(df_main, ["barcode"])
m_sku = find_col(df_main, ["sku"])
m_material = find_col(df_main, ["material"])
m_mgrp = find_col(df_main, ["mgrp"])


# ALL FILES KEYS
a_barcode = find_col(df_all, ["barcode"])
a_sku = find_col(df_all, ["sku", "item_number"])
a_material = find_col(df_all, ["material"])
a_mgrp = find_col(df_all, ["mgrp"])


print("Detected columns:")
print("MAIN:", m_barcode, m_sku, m_material, m_mgrp)
print("ALL :", a_barcode, a_sku, a_material, a_mgrp)


# ==============================
# RESULT DF
# ==============================

result = df_main.copy()

result["match_found"] = False
result["match_type"] = None


# ==============================
# SAFE MAP MATCH (NO MERGE CRASH)
# ==============================

def safe_match(main_col, all_col, name):

    global result

    if main_col is None or all_col is None:
        return

    print(f"Matching using {name}")

    # remove duplicates to avoid wrong mapping
    df_lookup = df_all[[all_col]].drop_duplicates()

    # create mapping series (SAFE)
    mapping = pd.Series(df_lookup[all_col].values, index=df_lookup[all_col])

    # apply map
    match_values = result[main_col].map(mapping)

    match_mask = match_values.notna()

    result[name + "_match"] = match_values

    result["match_found"] = result["match_found"] | match_mask

    result["match_type"] = result["match_type"].combine_first(
        match_mask.map(lambda x: name if x else None)
    )


# ==============================
# PRIORITY MATCHING
# ==============================

print("Matching started...")

safe_match(m_barcode, a_barcode, "barcode")
safe_match(m_sku, a_sku, "sku")
safe_match(m_material, a_material, "material")
safe_match(m_mgrp, a_mgrp, "mgrp")

print("Matching completed!")


# ==============================
# SAVE OUTPUT
# ==============================

result.to_csv(OUTPUT_FILE, index=False)

print(f"Saved → {OUTPUT_FILE}")