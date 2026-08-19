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

# ── SETTINGS ──────────────────────────
THRESHOLD      = 80   # name similarity threshold
NAME_CANDIDATES = 10   # top name candidates to check for weight match
SAMPLE_LIMIT   = 100    # stop as soon as this many matches are found

OUTPUT_FILE = r"C:\Users\HP\Desktop\fuzzy_sample_50.xlsx"


# ── FILE READER ───────────────────────

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
            print(f"   ❌  Cannot read {fname}.")
            return pd.DataFrame()

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


# ── PARSE DESCRIPTION → name, pack, weight ────────────────────────────

WEIGHT_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*(g|gm|gms|gram|grams|kg|ml|l|ltr|litre|liter|oz|cl)\b',
    re.IGNORECASE
)
PACK_RE = re.compile(
    r'(\d+\s*x\s*(?:\d+\s*x\s*)*)',
    re.IGNORECASE
)


def parse_description(text):
    if pd.isna(text):
        return "", "", ""
    s = str(text).strip().lower()

    weight = ""
    weight_matches = list(WEIGHT_RE.finditer(s))
    if weight_matches:
        wm = weight_matches[-1]
        weight = wm.group(0).replace(" ", "")
        s = s[:wm.start()] + " " + s[wm.end():]

    pack = ""
    pack_match = PACK_RE.search(s)
    if pack_match:
        pack = pack_match.group(0).replace(" ", "")
        s = s[:pack_match.start()] + " " + s[pack_match.end():]

    name = re.sub(r'\s+', ' ', s).strip()
    return name, pack, weight


def normalize_weight(w):
    if not w:
        return None
    m = re.match(r'(\d+(?:\.\d+)?)(g|gm|gms|gram|grams|kg|ml|l|ltr|litre|liter|oz|cl)',
                 w, re.IGNORECASE)
    if not m:
        return None
    val  = float(m.group(1))
    unit = m.group(2).lower()
    if unit in ("g", "gm", "gms", "gram", "grams"): return val
    if unit == "kg":                                  return val * 1000
    if unit == "ml":                                  return val
    if unit in ("l", "ltr", "litre", "liter"):        return val * 1000
    if unit == "cl":                                   return val * 10
    if unit == "oz":                                   return round(val * 28.3495, 2)
    return None


# ── FUZZY MATCH — stops at SAMPLE_LIMIT ──────────────────────────────

def fuzzy_match_sample(our_df, cust_df):
    matched = []

    # Pre-parse all our descriptions once
    our_parsed = []
    for _, row in our_df.iterrows():
        raw = str(row.get(OUR_COL, ""))
        name, pack, weight = parse_description(raw)
        our_parsed.append({
            "raw":      raw.strip().lower(),
            "name":     name,
            "pack":     pack,
            "weight":   weight,
            "weight_n": normalize_weight(weight),
        })

    our_names = [p["name"] for p in our_parsed]
    total     = len(cust_df)

    print(f"   Scanning customer rows — will stop at {SAMPLE_LIMIT} matches...\n")

    for i, cust_row in cust_df.iterrows():
        # ── Stop early once sample limit reached ──────────────────────
        if len(matched) >= SAMPLE_LIMIT:
            print(f"\n   ✅ Sample limit of {SAMPLE_LIMIT} reached at row {i} of {total}. Stopping.")
            break

        cust_raw = str(cust_row.get(CUST_COL, "")).strip().lower()
        if not cust_raw:
            continue

        cust_name, cust_pack, cust_weight = parse_description(cust_raw)
        cust_weight_n = normalize_weight(cust_weight)

        if not cust_name:
            continue

        # Step 1: top NAME candidates above threshold
        hits = process.extract(
            cust_name, our_names,
            scorer=fuzz.token_sort_ratio,
            limit=NAME_CANDIDATES,
            score_cutoff=THRESHOLD,
        )
        if not hits:
            continue

        # Step 2: find candidate where weight also matches
        best = None
        for hit_name, name_score, idx in hits:
            our = our_parsed[idx]

            weight_ok = False
            if our["weight_n"] is None and cust_weight_n is None:
                weight_ok = True
            elif our["weight_n"] is not None and cust_weight_n is not None:
                weight_ok = (our["weight_n"] == cust_weight_n)

            if weight_ok:
                if best is None or name_score > best["name_score"]:
                    best = {"idx": idx, "name_score": name_score}

        if best is None:
            continue

        our = our_parsed[best["idx"]]
        matched.append({
            "Our Full Desc":  our["raw"],
            "Our Name":       our["name"],
            "Our Pack":       our["pack"],
            "Our Weight":     our["weight"],
            "Cust Full Desc": cust_raw,
            "Cust Name":      cust_name,
            "Cust Pack":      cust_pack,
            "Cust Weight":    cust_weight,
            "Name Score":     round(best["name_score"], 1),
        })

        print(f"   [{len(matched):>2}/{SAMPLE_LIMIT}]  "
              f"'{our['raw'][:40]}'  ↔  '{cust_raw[:40]}'  "
              f"(score: {round(best['name_score'],1)}%)")

    return pd.DataFrame(matched)


# ── MAIN ──────────────────────────────

print("=" * 60)
print(f"  FUZZY SAMPLE — first {SAMPLE_LIMIT} matches")
print("=" * 60)
print(f"  Threshold      : {THRESHOLD}%")
print(f"  Our col        : {OUR_COL}")
print(f"  Customer col   : {CUST_COL}")
print(f"  Sample limit   : {SAMPLE_LIMIT} matches\n")

print("Loading our file...")
our_df = read_file(OUR_FILE)
print(f"   → {len(our_df):,} rows\n")

print("Loading customer file...")
cust_df = read_file(CUST_FILE)
print(f"   → {len(cust_df):,} rows\n")

print("Matching...\n")
matched_df = fuzzy_match_sample(our_df, cust_df)

print(f"\n   Total matched: {len(matched_df)}")

if not matched_df.empty:
    matched_df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"✅ Saved: {OUTPUT_FILE}")
else:
    print("⚠️  No matches found. Try lowering THRESHOLD.")