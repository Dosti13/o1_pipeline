import os, re, warnings
from pathlib import Path
import pandas as pd
from rapidfuzz import fuzz, process

warnings.filterwarnings("ignore")

# ── FILES ─────────────────────────────
OUR_FILE  = r"C:\Users\HP\Desktop\functionand sp.csv"
CUST_FILE = r"C:\Users\HP\Downloads\fwqniemtselloutdashboardsession2\C4 sep 24.xlsx"

OUR_COL  = "material_desc"
CUST_COL = "Item Name"

# ── THRESHOLD ─────────────────────────
THRESHOLD = 80

OUTPUT_FILE = r"C:\Users\HP\Desktop\fuzzy_result_c4.xlsx"


# ── FILE READER (from material_matcher) ───────────────────────────────

def detect_engine(filepath):
    ext = Path(filepath).suffix.lower()
    if ext in (".csv", ".tsv"):
        return "csv", ext
    with open(filepath, "rb") as f:
        header = f.read(8)
    if header[:4] == b"PK\x03\x04":
        return "pyxlsb" if ext == ".xlsb" else "openpyxl", ext
    if header[:4] == b"\xd0\xcf\x11\xe0":
        return "xlrd", ext
    with open(filepath, "rb") as f:
        head = f.read(512).lower()
    if b"<html" in head or b"<table" in head:
        return "html", ext
    return "openpyxl", ext


def read_file(filepath):
    engine, ext = detect_engine(filepath)
    fname       = Path(filepath).name
    frames      = []
    try:
        if engine == "csv":
            df = pd.read_csv(filepath, dtype=str,
                             sep="\t" if ext == ".tsv" else ",")
            df.dropna(how="all", inplace=True)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        if engine == "html":
            for i, df in enumerate(pd.read_html(filepath, dtype=str)):
                df.dropna(how="all", inplace=True)
                df.columns = [str(c).strip() for c in df.columns]
                frames.append(df)
        else:
            xls = pd.ExcelFile(filepath, engine=engine)
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
                df.dropna(how="all", inplace=True)
                df.columns = [str(c).strip() for c in df.columns]
                frames.append(df)
    except Exception as e:
        print(f"   ⚠️  Error reading {fname}: {e} — trying fallbacks...")
        for fallback in ["openpyxl", "xlrd"]:
            try:
                xls = pd.ExcelFile(filepath, engine=fallback)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
                    df.dropna(how="all", inplace=True)
                    df.columns = [str(c).strip() for c in df.columns]
                    frames.append(df)
                print(f"   ✅  Loaded with '{fallback}'")
                break
            except Exception:
                continue
        else:
            print(f"   ❌  Cannot read {fname}. Skipping.")
            return pd.DataFrame()

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


# ── FUZZY MATCH ───────────────────────

def fuzzy_match(our_df, cust_df):
    matched = []
    our_texts = our_df[OUR_COL].fillna("").astype(str).str.strip().str.lower().tolist()
    total = len(cust_df)

    for i, cust_row in cust_df.iterrows():
        cust_val = str(cust_row.get(CUST_COL, "")).strip().lower()
        if not cust_val:
            continue

        if (i + 1) % 200 == 0:
            print(f"   {i + 1}/{total}...")

        hit = process.extractOne(
            cust_val, our_texts,
            scorer=fuzz.token_sort_ratio,
            score_cutoff=THRESHOLD,
        )

        if hit:
            best_val, score, idx = hit
            matched.append({
                "Our Description":  best_val,
                "Customer Product": cust_val,
                "Score":            round(score, 1),
            })

    return pd.DataFrame(matched)


# ── MAIN ──────────────────────────────

print(f"Threshold: {THRESHOLD}%")

print("Loading our file...")
our_df = read_file(OUR_FILE)
print(f"   → {len(our_df):,} rows")

print("Loading customer file...")
cust_df = read_file(CUST_FILE)
print(f"   → {len(cust_df):,} rows")

print("Matching...")
matched_df = fuzzy_match(our_df, cust_df)

print(f"\n   Matched: {len(matched_df):,}")

matched_df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")

print(f"\n✅ Done! Saved: {OUTPUT_FILE}")