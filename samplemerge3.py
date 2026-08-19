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
THRESHOLD       = 80    # name similarity threshold
NAME_CANDIDATES = 10    # top name candidates to check
SAMPLE_LIMIT    = 100   # stop after this many matches

OUTPUT_FILE = r"C:\Users\HP\Desktop\fuzzy_sample_10.xlsx"


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
    """Split a product description into (name, pack, weight)."""
    if pd.isna(text):
        return "", "", ""
    s = str(text).strip().lower()

    # Extract weight (take the LAST match — usually the unit weight)
    weight = ""
    weight_matches = list(WEIGHT_RE.finditer(s))
    if weight_matches:
        wm = weight_matches[-1]
        weight = wm.group(0).replace(" ", "")
        s = s[:wm.start()] + " " + s[wm.end():]

    # Extract pack pattern like 4x6x, 12x, 2x3x etc.
    pack = ""
    pack_match = PACK_RE.search(s)
    if pack_match:
        pack = pack_match.group(0).replace(" ", "").rstrip("x").lower()
        s = s[:pack_match.start()] + " " + s[pack_match.end():]

    # Remaining text is the product name
    name = re.sub(r'\s+', ' ', s).strip()
    return name, pack, weight


def normalize_weight(w):
    """Convert any weight string to a common unit (grams or ml)."""
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


def normalize_pack(p):
    """Normalize pack string for comparison: '4x6x' → '4x6'."""
    if not p:
        return None
    # Strip trailing 'x', normalize spaces
    clean = p.strip().lower().rstrip("x")
    return clean if clean else None


# ── FUZZY MATCH — Name → Weight → Pack priority ─────────────────────

def fuzzy_match_sample(our_df, cust_df):
    matched = []

    # ── Step 0: Pre-parse ALL our descriptions once ──────────────────
    our_parsed = []
    for _, row in our_df.iterrows():
        raw = str(row.get(OUR_COL, ""))
        name, pack, weight = parse_description(raw)
        our_parsed.append({
            "raw":      raw.strip().lower(),
            "name":     name,
            "pack":     normalize_pack(pack),
            "pack_raw": pack,
            "weight":   weight,
            "weight_n": normalize_weight(weight),
        })

    our_names = [p["name"] for p in our_parsed]
    total     = len(cust_df)

    print(f"   Scanning {total:,} customer rows — stop at {SAMPLE_LIMIT} matches...\n")

    for i, cust_row in cust_df.iterrows():
        if len(matched) >= SAMPLE_LIMIT:
            print(f"\n   ✅ Sample limit of {SAMPLE_LIMIT} reached at row {i}. Stopping.")
            break

        cust_raw = str(cust_row.get(CUST_COL, "")).strip().lower()
        if not cust_raw:
            continue

        cust_name, cust_pack, cust_weight = parse_description(cust_raw)
        cust_weight_n = normalize_weight(cust_weight)
        cust_pack_n   = normalize_pack(cust_pack)

        if not cust_name:
            continue

        # ── STEP 1: Fuzzy match on NAME only ─────────────────────────
        hits = process.extract(
            cust_name, our_names,
            scorer=fuzz.token_sort_ratio,
            limit=NAME_CANDIDATES,
            score_cutoff=THRESHOLD,
        )
        if not hits:
            continue

        # ── STEP 2: Among name hits, filter by WEIGHT match ─────────
        weight_matched = []
        for hit_name, name_score, idx in hits:
            our = our_parsed[idx]
            weight_ok = False

            if our["weight_n"] is None and cust_weight_n is None:
                weight_ok = True          # both have no weight → ok
            elif our["weight_n"] is not None and cust_weight_n is not None:
                weight_ok = (our["weight_n"] == cust_weight_n)
            # If one has weight and the other doesn't → NOT a match

            if weight_ok:
                weight_matched.append((idx, name_score))

        if not weight_matched:
            continue   # no candidate with matching weight

        # ── STEP 3: Among weight-matched, prefer PACK match ─────────
        pack_matched = []
        pack_unmatched = []

        for idx, name_score in weight_matched:
            our = our_parsed[idx]
            pack_ok = False

            if our["pack"] is None and cust_pack_n is None:
                pack_ok = True            # both have no pack → ok
            elif our["pack"] is not None and cust_pack_n is not None:
                pack_ok = (our["pack"] == cust_pack_n)
            # If one has pack and the other doesn't → not a pack match

            if pack_ok:
                pack_matched.append((idx, name_score))
            else:
                pack_unmatched.append((idx, name_score))

        # Pick from pack-matched first; fall back to weight-only match
        candidates = pack_matched if pack_matched else pack_unmatched
        # Among final candidates, pick the highest name score
        best_idx, best_score = max(candidates, key=lambda x: x[1])

        our = our_parsed[best_idx]
        match_level = "Name+Weight+Pack" if pack_matched else "Name+Weight"

        matched.append({
            "Our Full Desc":   our["raw"],
            "Our Name":        our["name"],
            "Our Pack":        our["pack_raw"],
            "Our Weight":      our["weight"],
            "Cust Full Desc":  cust_raw,
            "Cust Name":       cust_name,
            "Cust Pack":       cust_pack,
            "Cust Weight":     cust_weight,
            "Name Score":      round(best_score, 1),
            "Match Level":     match_level,
        })

        print(f"   [{len(matched):>3}/{SAMPLE_LIMIT}]  "
              f"'{our['raw'][:42]}'  ↔  '{cust_raw[:42]}'  "
              f"(score: {best_score:.0f}% | {match_level})")

    return pd.DataFrame(matched)


# ── MAIN ──────────────────────────────

print("=" * 60)
print(f"  FUZZY MATCH v2 — Name → Weight → Pack")
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

print(f"\n   Total matched : {len(matched_df)}")
if not matched_df.empty:
    # Summary by match level
    print(f"   Name+Weight+Pack : {(matched_df['Match Level'] == 'Name+Weight+Pack').sum()}")
    print(f"   Name+Weight only : {(matched_df['Match Level'] == 'Name+Weight').sum()}")
    matched_df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"\n✅ Saved: {OUTPUT_FILE}")
else:
    print("⚠️  No matches found. Try lowering THRESHOLD.")