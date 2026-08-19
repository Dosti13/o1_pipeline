"""
Text-first matcher v2 — NAME + WEIGHT split matching
======================================================
Changes vs v1:
  1. Each description is split into (name_part, weight_part).
  2. Matching is two-stage:
       a. Name similarity (token_set_ratio on name_part) — primary score.
       b. Weight must match within 5 % tolerance — hard filter (score → 0 if mismatch).
  3. Brand is now MANDATORY: if brands don't overlap, the match is dropped entirely.
  4. Pack format (e.g. 4X12, 6X3) is normalised and compared as a soft bonus (+3).

Configure the CONFIGURATION block below, then run:
    python text_matcher_v2.py

To see how sample descriptions are parsed without any files:
    python text_matcher_v2.py --demo
"""

# ===========================================================================
# CONFIGURATION — edit these values before running
# ===========================================================================

OUR_FILE    = r"C:\Users\HP\Desktop\functionand sp.csv"

CUST_FILE   = r"C:\Users\HP\Downloads\fwqniemtselloutdashboardsession2\C4 sep 24.xlsx"
CUST_LABEL  = "C4"
CUST_READER = "read_c4"      # read_almeera | read_c4 | read_grandmall |
                              # read_qnie | read_rawspar | read_talabat

OUTPUT_FILE = r"C:\Users\HP\Desktop\match_results_v.xlsx"

SAMPLE_SIZE    = 300   # e.g. 200  or  None to keep all rows
MIN_CONFIDENCE = 70     # rows below this are dropped

# ===========================================================================

import pandas as pd
import numpy as np
import re
import time
from collections import defaultdict
from rapidfuzz import fuzz, process

# ---------------------------------------------------------------------------
# Abbreviation tables
# ---------------------------------------------------------------------------
ABBREV = {
    r'\bPKT\b': 'PACKET', r'\bPKTS\b': 'PACKET', r'\bBX\b': 'BOX',
    r'\bBTL\b': 'BOTTLE', r'\bSACH\b': 'SACHET', r'\bSACHE\b': 'SACHET',
    r'\bPC\b': 'PIECE', r'\bPCS\b': 'PIECE',
    r'\bFZ\b': 'FROZEN', r'\bFRZN\b': 'FROZEN', r'\bFRSH\b': 'FRESH',
    r'\bDRD\b': 'DRIED', r'\bRSTD\b': 'ROASTED', r'\bSMKD\b': 'SMOKED',
    r'\bGRLLD\b': 'GRILLED', r'\bMRND\b': 'MARINATED',
    r'\bINST\b': 'INSTANT', r'\bACTV\b': 'ACTIVE', r'\bORG\b': 'ORGANIC',
    r'\bORGNC\b': 'ORGANIC', r'\bNAT\b': 'NATURAL', r'\bNTRL\b': 'NATURAL',
    r'\bUNSWT\b': 'UNSWEETENED', r'\bUNSWTND\b': 'UNSWEETENED',
    r'\bSWT\b': 'SWEET', r'\bUNSL?TD\b': 'UNSALTED', r'\bSLTD\b': 'SALTED',
    r'\bPWDR?\b': 'POWDER', r'\bPWR\b': 'POWDER', r'\bBKG\b': 'BAKING',
    r'\bBAK\b': 'BAKING', r'\bBICRBNTE\b': 'BICARBONATE',
    r'\bBICARB\b': 'BICARBONATE',
    r'\bCKN\b': 'CHICKEN', r'\bCHKN\b': 'CHICKEN', r'\bCHCKN\b': 'CHICKEN',
    r'\bCHS\b': 'CHEESE', r'\bCHEZ\b': 'CHEESE', r'\bCHDR\b': 'CHEDDAR',
    r'\bMZRLA\b': 'MOZZARELLA', r'\bMZRELLA\b': 'MOZZARELLA',
    r'\bVEG\b': 'VEGETABLE', r'\bVGE\b': 'VEGETABLE',
    r'\bGHE\b': 'GHEE', r'\bBTR\b': 'BUTTER', r'\bBTTR\b': 'BUTTER',
    r'\bCHOC\b': 'CHOCOLATE', r'\bCHOCO\b': 'CHOCOLATE',
    r'\bSTRWB(?:RY|ERY|ERRY)?\b': 'STRAWBERRY',
    r'\bRSPB(?:RY|ERY|ERRY)?\b': 'RASPBERRY',
    r'\bBLBR(?:RY|ERY|ERRY)?\b': 'BLUEBERRY',
    r'\bEVDAY\b': 'EVERYDAY',
    r'\bHNY\b': 'HONEY',
    r'\bBISC\b': 'BISCUIT',
    r'\bGLUC\b': 'GLUCOSE',
    r'\bPNAPL\b': 'PINEAPPLE', r'\bPNPL\b': 'PINEAPPLE',
    r'\bMNGO\b': 'MANGO', r'\bAPL\b': 'APPLE',
    r'\bBRZ\b': 'BRAZIL', r'\bBRZL\b': 'BRAZIL',
}

BRAND_ABBREV = {
    'AME': 'AMERICANA', 'AMC': 'AMERICANA', 'AMER': 'AMERICANA',
    'AMRCNA': 'AMERICANA', 'TIFF': 'TIFFANY', 'TIFFFANY': 'TIFFANY',
    'DEEM': 'DEEMAH', 'AMT': 'AHMADTEA', 'AHMAD': 'AHMADTEA',
    'AHMADTEA': 'AHMADTEA', 'FOSTER': 'FOSTERCLARK',
    'FOSTERCLARK': 'FOSTERCLARK', 'FOSTERCLARKS': 'FOSTERCLARK',
    'GG': 'GREENGIANT', 'GREENGIANT': 'GREENGIANT',
    'HAPPYCOW': 'HAPPYCOW', 'NV': 'NATUREVALLEY',
    'NATUREVALLEY': 'NATUREVALLEY', 'CRF': 'CARREFOUR',
    'MCRF': 'CARREFOUR', 'PRES': 'PRESIDENT', 'HERSHEYS': 'HERSHEY',
    'HERSHY': 'HERSHEY', 'BC': 'BETTYCROCKER', 'BETTY': 'BETTYCROCKER',
    'BETTYCROCKER': 'BETTYCROCKER', 'LD': 'LULUDAILY',
    'LULUDAILY': 'LULUDAILY',
}

NON_BRAND = {
    'FRESH', 'FROZEN', 'PREMIUM', 'NATURAL', 'ORGANIC', 'PURE', 'WHOLE',
    'GRILLED', 'ROASTED', 'SMOKED', 'BAKED', 'CHILLED', 'DRIED', 'CANNED',
    'HOT', 'COLD', 'MINI', 'MAXI', 'SUPER', 'EXTRA', 'LIGHT', 'GREEN',
    'RED', 'WHITE', 'BLUE', 'BLACK', 'YELLOW', 'BROWN', 'BIG', 'SMALL',
    'MED', 'NEW', 'OLD', 'CV', 'EA', 'KG', 'GR', 'GM', 'ML',
    'CH', 'BRZ', 'BRAZIL', 'AMERICAN', 'SF', 'BK', 'F', 'FV',
    'BREADED', 'JUMBO', 'SLICED', 'SLICE', 'STUFFED', 'STRIPS',
    'NUGGETS', 'CHICKEN', 'BEEF', 'FISH', 'CHEESE', 'TEA', 'COFFEE',
    'JUICE', 'WATER', 'OIL', 'RICE', 'PASTA', 'LOCAL', 'BRAND',
    'NONE', 'NA', 'NAN',
    'CHOCOLATE', 'CHOCO', 'VANILLA', 'STRAWBERRY', 'RASPBERRY',
    'BLUEBERRY', 'ORANGE', 'LEMON', 'MANGO', 'PINEAPPLE', 'BANANA',
    'APPLE', 'PEACH', 'CHERRY', 'MINT', 'CARAMEL', 'COCOA', 'HONEY',
    'PLAIN', 'SPICY', 'SWEET', 'SOUR', 'SALT', 'PEPPER',
    'FLOUR', 'SUGAR', 'MILK', 'CREAM', 'YOGURT', 'BUTTER', 'GHEE',
    'BREAD', 'CAKE', 'COOKIES', 'BISCUIT', 'BISCUITS', 'WAFER',
    'CRISPY', 'CRUNCHY', 'DELUXE', 'CLASSIC', 'STANDARD', 'REGULAR',
    'FULL', 'HALF', 'QUARTER', 'WS',
}

# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------

# Pack notation: 4X12X50G, 6X3X250G, 1X24X230GR, 24X84G, 4X12X50
# We want to capture this as a block to strip from the name.
PACK_RE = re.compile(
    r'\b(\d+\s*[Xx]\s*\d+(?:\s*[Xx]\s*\d+(?:\.\d+)?)?'
    r'(?:\s*(?:KG|KGS|G|GR|GM|GMS|ML|L|LT|LTR|LTRS|MG))?)\b',
    re.IGNORECASE,
)

SIZE_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*'
    r'(KG|KGS|G|GR|GM|GMS|GRM|GRMS|GRAMS?|MG|'
    r'L|LT|LTR|LTRS|LITRE|LITER|ML|CL|CC)\b',
    re.IGNORECASE,
)

UNIT_TABLE = {
    'KG': (1000, 'mass'), 'KGS': (1000, 'mass'),
    'G': (1, 'mass'), 'GR': (1, 'mass'), 'GM': (1, 'mass'),
    'GMS': (1, 'mass'), 'GRM': (1, 'mass'), 'GRMS': (1, 'mass'),
    'GRAM': (1, 'mass'), 'GRAMS': (1, 'mass'), 'MG': (0.001, 'mass'),
    'L': (1000, 'vol'), 'LT': (1000, 'vol'), 'LTR': (1000, 'vol'),
    'LTRS': (1000, 'vol'), 'LITRE': (1000, 'vol'), 'LITER': (1000, 'vol'),
    'ML': (1, 'vol'), 'CL': (10, 'vol'), 'CC': (1, 'vol'),
}

TOKEN_RE = re.compile(r"[A-Z0-9&]+")

# ---------------------------------------------------------------------------
# Core parser: split description → (name, weight_g_or_ml, kind, pack_str)
# ---------------------------------------------------------------------------

def parse_description(text: str):
    """
    Split a product description into:
      name       – product name with all size/pack tokens removed
      weight     – per-unit weight in grams or ml (smallest size token found)
      kind       – 'mass' | 'vol' | None
      pack_str   – normalised outer pack e.g. '4X12', '24', '6X3'

    Examples:
      'TIFFANY DIGESTIVE NATURAL 6X3X250G'
          → name='TIFFANY DIGESTIVE NATURAL', weight=250g, kind='mass', pack='6X3'
      'TIFF NICE EVDAY REG 4X12X50G'
          → name='TIFFANY NICE EVERYDAY REG', weight=50g, kind='mass', pack='4X12'
      'ROYAL CLASSIC BUTTER COOKIES 454GX12'
          → name='ROYAL CLASSIC BUTTER COOKIES', weight=454g, kind='mass', pack='12'
    """
    if not isinstance(text, str):
        return ('', None, None, '')

    s = text.upper()

    # Expand abbreviations
    s = re.sub(r'[=\.,;:]', ' ', s)
    for pat, rep in ABBREV.items():
        s = re.sub(pat, rep, s)

    # Collect all size tokens (number + unit)
    sizes_found = []
    for m in SIZE_RE.finditer(s):
        unit = m.group(2).upper()
        if unit not in UNIT_TABLE:
            continue
        mult, kind = UNIT_TABLE[unit]
        try:
            sizes_found.append((float(m.group(1)) * mult, kind))
        except ValueError:
            pass

    # Also catch "454GX12" and "135MLX10" patterns: size glued directly to XN
    GLUED_RE = re.compile(
        r'(\d+(?:\.\d+)?)\s*(KG|KGS|G|GR|GM|GMS|ML|L|LT|LTR|MG)\s*X\s*(\d+)\b',
        re.IGNORECASE,
    )
    for m in GLUED_RE.finditer(s):
        unit = m.group(2).upper()
        if unit not in UNIT_TABLE:
            continue
        mult, kind = UNIT_TABLE[unit]
        try:
            sizes_found.append((float(m.group(1)) * mult, kind))
        except ValueError:
            pass

    # Per-unit weight = smallest size value found
    weight_base = None
    weight_kind = None
    if sizes_found:
        sizes_found.sort(key=lambda x: x[0])
        weight_base, weight_kind = sizes_found[0]

    # Find pack tokens and extract outer count (everything except the innermost unit)
    pack_str = ''
    pack_tokens = list(PACK_RE.finditer(s))
    GLUED_PACK_RE = re.compile(
        r'(\d+(?:\.\d+)?)\s*(KG|KGS|G|GR|GM|GMS|ML|L|LT|LTR|MG)\s*X\s*(\d+)\b',
        re.IGNORECASE,
    )
    if pack_tokens:
        best_pack = max(pack_tokens, key=lambda m: len(m.group(0)))
        raw = best_pack.group(0).upper()
        raw = SIZE_RE.sub('', raw).strip()          # strip unit suffix
        raw = re.sub(r'\s*[Xx]\s*', 'X', raw).strip('X')
        pack_str = raw
    else:
        # Handle "454GX12" / "135MLX10" style
        gm = GLUED_PACK_RE.search(s)
        if gm:
            pack_str = gm.group(3)

    # Build clean name: strip glued, pack and size tokens
    name = s
    name = GLUED_PACK_RE.sub(' ', name)             # 454GX12 → remove whole thing
    name = PACK_RE.sub(' ', name)
    name = SIZE_RE.sub(' ', name)
    name = re.sub(r'\bX\s*\d+\b', ' ', name)       # stray "X 12" etc.
    name = re.sub(r'\bWS\b', ' ', name)             # "WS" suffix
    name = re.sub(r'\s+', ' ', name).strip()

    return (name, weight_base, weight_kind, pack_str)


# ---------------------------------------------------------------------------
# Brand helpers
# ---------------------------------------------------------------------------

def normalise_brand_str(raw):
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return frozenset()
    s = str(raw).upper()
    s = s.replace("'", '').replace('-', ' ').replace('.', ' ').replace('/', ' ').replace(':', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    tokens = TOKEN_RE.findall(s)
    out = set()
    if len(tokens) >= 2:
        joined = ''.join(tokens[:3])
        if joined not in NON_BRAND and len(joined) >= 4:
            out.add(BRAND_ABBREV.get(joined, joined))
    for t in tokens:
        if t in NON_BRAND or t.isdigit() or len(t) < 2:
            continue
        out.add(BRAND_ABBREV.get(t, t))
    return frozenset(out)


def brand_tokens_from_desc(clean_name: str) -> frozenset:
    if not clean_name:
        return frozenset()
    words = clean_name.split()[:3]
    out = set()
    for w in words:
        if w in NON_BRAND or w.isdigit() or len(w) < 2:
            continue
        out.add(BRAND_ABBREV.get(w, w))
    return frozenset(out)


def brands_overlap(a: frozenset, b: frozenset) -> bool:
    if not a or not b:
        return False
    for x in a:
        for y in b:
            if x == y:
                return True
            if len(x) >= 4 and len(y) >= 4 and (x in y or y in x):
                return True
            if len(x) >= 4 and len(y) >= 4 and x[:4] == y[:4]:
                return True
    return False


# ---------------------------------------------------------------------------
# Size band helper (for candidate pre-filter)
# ---------------------------------------------------------------------------

def size_band(size):
    if size is None or (isinstance(size, float) and np.isnan(size)) or size <= 0:
        return None
    return int(round(np.log10(size) * 20))


def weights_match(w1, w2, tol=0.05):
    """
    Returns:
      True  — weights are within tol of each other
      False — weights are present but don't match
      None  — one or both weights are unknown (can't decide)
    """
    if w1 is None or w2 is None:
        return None
    if (isinstance(w1, float) and np.isnan(w1)) or (isinstance(w2, float) and np.isnan(w2)):
        return None
    if w1 <= 0 or w2 <= 0:
        return None
    return (min(w1, w2) / max(w1, w2)) >= (1.0 - tol)


def pack_match_bonus(p1: str, p2: str) -> float:
    return 3.0 if (p1 and p2 and p1 == p2) else 0.0


# ---------------------------------------------------------------------------
# Customer file readers
# ---------------------------------------------------------------------------

def read_almeera(path):
    df = pd.read_excel(path, usecols=['Product name', 'Brand'])
    df = df.dropna(subset=['Product name']).drop_duplicates(subset=['Product name']).reset_index(drop=True)
    return df.rename(columns={'Product name': 'desc', 'Brand': 'brand'})

def read_c4(path):
    df = pd.read_excel(path, sheet_name='Items', usecols=['Item Name', 'Brand'])
    df = df.dropna(subset=['Item Name']).drop_duplicates(subset=['Item Name']).reset_index(drop=True)
    return df.rename(columns={'Item Name': 'desc', 'Brand': 'brand'})

def read_grandmall(path):
    df = pd.read_excel(path, usecols=['SU_DESCRIPTION', 'BRAND'])
    df = df.dropna(subset=['SU_DESCRIPTION']).drop_duplicates(subset=['SU_DESCRIPTION']).reset_index(drop=True)
    return df.rename(columns={'SU_DESCRIPTION': 'desc', 'BRAND': 'brand'})

def read_qnie(path):
    df = pd.read_excel(path, header=3)
    desc_col = df.columns[4]
    df = df[['Brand', desc_col]].copy()
    df.columns = ['brand', 'desc']
    df = df.dropna(subset=['desc']).drop_duplicates(subset=['desc']).reset_index(drop=True)
    return df[['desc', 'brand']]

def read_rawspar(path):
    df = pd.read_excel(path, usecols=['Family Text', 'Retail Article Brand Description Text'])
    df = df.rename(columns={
        'Family Text': 'family',
        'Retail Article Brand Description Text': 'brand'
    })
    df = df.dropna(subset=['brand']).drop_duplicates(subset=['family', 'brand']).reset_index(drop=True)
    df['desc'] = df['brand'].astype(str).str.upper() + ' ' + df['family'].astype(str).str.upper()
    return df[['desc', 'brand']]

def read_talabat(path):
    df = pd.read_excel(path, usecols=['Product_Name', 'Brand'])
    df = df.dropna(subset=['Product_Name']).drop_duplicates(subset=['Product_Name']).reset_index(drop=True)
    return df.rename(columns={'Product_Name': 'desc', 'Brand': 'brand'})

READERS = {
    'read_almeera':   read_almeera,
    'read_c4':        read_c4,
    'read_grandmall': read_grandmall,
    'read_qnie':      read_qnie,
    'read_rawspar':   read_rawspar,
    'read_talabat':   read_talabat,
}


# ===========================================================================
# Main
# ===========================================================================

def main():
    # ------------------------------------------------------------------ #
    # 1. Load & parse source catalogue
    # ------------------------------------------------------------------ #
    print("Loading source catalogue...")
    src = pd.read_csv(OUR_FILE, low_memory=False)
    src = src.dropna(subset=['material_desc']).reset_index(drop=True)

    print("  Parsing source descriptions (name + weight)...")
    parsed_src = [parse_description(t) for t in src['material_desc']]
    src['_name']   = [p[0] for p in parsed_src]
    src['_weight'] = [p[1] for p in parsed_src]
    src['_kind']   = [p[2] for p in parsed_src]
    src['_pack']   = [p[3] for p in parsed_src]
    src['_band']   = src['_weight'].apply(size_band)

    src['_brands'] = [
        normalise_brand_str(mg) | brand_tokens_from_desc(name)
        for mg, name in zip(src['mgrp_descr'], src['_name'])
    ]
    print(f"  Source rows: {len(src):,}")

    # Size-band index for fast candidate lookup
    src_buckets: dict = defaultdict(list)
    for i, (kind, band) in enumerate(zip(src['_kind'], src['_band'])):
        kind = None if (isinstance(kind, float) and pd.isna(kind)) else kind
        src_buckets[(kind, band)].append(i)

    # ------------------------------------------------------------------ #
    # 2. Load & parse customer file
    # ------------------------------------------------------------------ #
    reader_fn = READERS.get(CUST_READER)
    if reader_fn is None:
        raise ValueError(f"Unknown reader '{CUST_READER}'. Options: {list(READERS)}")

    print(f"\nLoading customer file ({CUST_LABEL})...")
    cust = reader_fn(CUST_FILE)
    cust['source_file'] = CUST_LABEL
    print(f"  Customer rows: {len(cust):,}")

    print("  Parsing customer descriptions (name + weight)...")
    parsed_cust = [parse_description(t) for t in cust['desc']]
    cust['_name']   = [p[0] for p in parsed_cust]
    cust['_weight'] = [p[1] for p in parsed_cust]
    cust['_kind']   = [p[2] for p in parsed_cust]
    cust['_pack']   = [p[3] for p in parsed_cust]
    cust['_band']   = cust['_weight'].apply(size_band)

    cust['_brands'] = [
        normalise_brand_str(b) | brand_tokens_from_desc(name)
        for b, name in zip(cust['brand'], cust['_name'])
    ]

    # ------------------------------------------------------------------ #
    # 3. Matching
    # ------------------------------------------------------------------ #
    print("\nMatching  [name similarity → weight hard filter → brand mandatory]...")
    results: list = [None] * len(cust)
    t0 = time.time()

    cust_groups: dict = defaultdict(list)
    for ci, (kind, band) in enumerate(zip(cust['_kind'], cust['_band'])):
        kind = None if (isinstance(kind, float) and pd.isna(kind)) else kind
        cust_groups[(kind, band)].append(ci)
    print(f"  {len(cust_groups):,} customer (kind, band) buckets")

    # Pre-extract as plain lists for speed
    src_name    = src['_name'].tolist()
    src_weight  = src['_weight'].tolist()
    src_kind    = src['_kind'].tolist()
    src_pack    = src['_pack'].tolist()
    src_brands  = src['_brands'].tolist()
    src_descraw = src['material_desc'].tolist()

    cust_name   = cust['_name'].tolist()
    cust_weight = cust['_weight'].tolist()
    cust_kind   = cust['_kind'].tolist()
    cust_pack   = cust['_pack'].tolist()
    cust_brands = cust['_brands'].tolist()

    bcount = 0
    for (kind, band), cidxs in cust_groups.items():
        bcount += 1

        # Candidate selection (size-band ± 1 bucket)
        if band is None:
            cand_idxs = list(range(len(src)))
        else:
            cand_idxs = []
            for d in (-1, 0, 1):
                cand_idxs.extend(src_buckets.get((kind, band + d), []))
            cand_idxs.extend(src_buckets.get((kind, None), []))
            cand_idxs.extend(src_buckets.get((None, None), []))

        if not cand_idxs:
            for ci in cidxs:
                results[ci] = (None, 0.0, 'no_candidates')
            continue

        cand_names  = [src_name[i]   for i in cand_idxs]
        query_names = [cust_name[ci] for ci in cidxs]

        if not query_names or not cand_names:
            for ci in cidxs:
                results[ci] = (None, 0.0, 'empty_names')
            continue

        # --- Step A: Name similarity (vectorised) ---
        mat = process.cdist(
            query_names, cand_names,
            scorer=fuzz.token_set_ratio, workers=-1, dtype=np.float32,
        )

        for q_i, ci in enumerate(cidxs):
            scores = mat[q_i].copy().astype(np.float64)

            c_weight = cust_weight[ci]
            c_kind_v = cust_kind[ci]
            c_pack_v = cust_pack[ci]
            c_brands = cust_brands[ci]

            # --- Step B: Weight hard filter + pack bonus ---
            for j, si in enumerate(cand_idxs):
                wm = weights_match(c_weight, src_weight[si])
                if wm is False:
                    # Hard exclude — wrong weight
                    scores[j] = 0.0
                    continue
                if wm is True:
                    scores[j] = min(100.0, scores[j] + 5.0)
                    # Penalise if mass vs vol clash
                    if c_kind_v and src_kind[si] and c_kind_v != src_kind[si]:
                        scores[j] = max(0.0, scores[j] - 20.0)
                # wm is None → unchanged
                # Pack bonus
                scores[j] = min(100.0, scores[j] + pack_match_bonus(c_pack_v, src_pack[si]))

            # --- Step C: Brand MANDATORY ---
            # Only consider candidates where brand overlaps.
            best_score  = 0.0
            best_src_i  = None
            best_reason = 'no_brand_match'

            for j, si in enumerate(cand_idxs):
                if scores[j] < MIN_CONFIDENCE:
                    continue
                if brands_overlap(c_brands, src_brands[si]):
                    if scores[j] > best_score:
                        best_score  = scores[j]
                        best_src_i  = si
                        best_reason = 'ok'

            if best_src_i is None:
                # Record best non-brand match for diagnostics only
                best_j  = int(np.argmax(scores))
                best_si = cand_idxs[best_j]
                results[ci] = (best_si, float(scores[best_j]), 'brand_mismatch')
            else:
                results[ci] = (best_src_i, best_score, best_reason)

        if bcount % 20 == 0 or bcount == len(cust_groups):
            print(f"  {bcount}/{len(cust_groups)}  ({time.time()-t0:.1f}s)")

    print(f"Matching done in {time.time()-t0:.1f}s")

    # ------------------------------------------------------------------ #
    # 4. Build output
    # ------------------------------------------------------------------ #
    print(f"\nBuilding output  (≥{MIN_CONFIDENCE} confidence, brand mandatory)...")
    all_rows = []

    for ci in range(len(cust)):
        result = results[ci]
        if result is None:
            continue
        src_i, conf, reason = result

        if src_i is None or reason == 'brand_mismatch':
            continue

        conf_int = int(round(conf))
        if conf_int < MIN_CONFIDENCE:
            continue

        status = 'Matched (High)' if conf_int >= 85 else 'Matched (Medium)'

        all_rows.append({
            'Source File':              cust.at[ci, 'source_file'],
            # Our side
            'Our Description':          src_descraw[src_i],
            'Our Name (parsed)':        src['_name'].iloc[src_i],
            'Our Weight g/ml':          src['_weight'].iloc[src_i],
            'Our Pack':                 src['_pack'].iloc[src_i],
            # Customer side
            'Customer Description':     cust.at[ci, 'desc'],
            'Customer Name (parsed)':   cust['_name'].iloc[ci],
            'Customer Weight g/ml':     cust['_weight'].iloc[ci],
            'Customer Pack':            cust['_pack'].iloc[ci],
            'Customer Brand':           cust.at[ci, 'brand'],
            # Result
            'Brand Check':              'OK',
            'Match Status':             status,
            'Confidence Score':         conf_int,
        })

    out = pd.DataFrame(all_rows)
    print(f"Total matched rows (≥{MIN_CONFIDENCE}, brand OK): {len(out):,}")

    if SAMPLE_SIZE and SAMPLE_SIZE > 0 and len(out) > SAMPLE_SIZE:
        out = out.head(SAMPLE_SIZE)
        print(f"Sample applied — keeping first {SAMPLE_SIZE} rows.")

    if not out.empty:
        print("\nMatch status breakdown:")
        print(out['Match Status'].value_counts().to_string())

    out.to_excel(OUTPUT_FILE, index=False)
    print(f"\nSaved → {OUTPUT_FILE}")
    print(f"Rows written: {len(out):,}")


# ===========================================================================
# --demo mode: show how sample data is parsed (no files needed)
# ===========================================================================

def demo_parse():
    samples = [
        "TIFFANY DIGESTIVE NATURAL 6X3X250G",
        "TIFF NICE EVDAY REG 4X12X50G",
        "TIFF CREAM EVDAY CHOCO 4X6X 90G",
        "TIFF CREAM EVDAY ORANGE 4X6X 90G",
        "TIFF CREAM EVDAY MANGO 4X6X 90G",
        "TIFF NICE EVDAY REG 4X12X50G",
        "TIFF GLUC MILK AND HNY BISC 4X12X40GR",
        "TIFF CREAM EVDAY STRWBRY 24X84G WS",
        "TIFF GLUCOSE MILK AND HONEY BISC 4X12X50",
        "ROYAL CLASSIC BUTTER COOKIES 454GX12",
        "OREO CHOCOLATE SANDWICH 135MLX10",
        "DEEMAH DIGESTIVE BISCUIT 1X24X230GR",
        "SANDWICH BISCUITS CHOCOLATE 90GR X 24",
        "TIFF NUTTY BITES CASHEW 24X81G",
        "CHOCO CHIP COOKIE 496GR X 12",
        "TIFF NUTTY BITES PISTACHIO 24X81G",
        "OREO CHOCOLATE SANDWICH 135MLX10",
    ]
    print(f"\n{'Original':<46} {'Name (parsed)':<40} {'Wt g/ml':>10}  {'Kind':>5}  Pack")
    print('-' * 115)
    for s in samples:
        name, wt, kind, pack = parse_description(s)
        wt_str = f"{wt:.1f}" if wt is not None else '-'
        print(f"{s:<46} {name:<40} {wt_str:>10}  {(kind or '-'):>5}  {pack}")


if __name__ == '__main__':
    import sys
    if '--demo' in sys.argv:
        demo_parse()
    else:
        main()