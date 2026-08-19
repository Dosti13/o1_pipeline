"""
Combined Matcher — BARCODE + NAME + WEIGHT + BRAND + HIERARCHY BONUS  (Barcode → Fuzzy → Semantic cascade)
=============================================================================================================
VERSION 5.2 — BRIDGE PROVENANCE + BORROWED BARCODE  (adapted from V5.1)
─────────────────────────────────────────────────────────────────────────────────────
SAME MATCHING LOGIC AS V5.1. What changed, and why:

  OBSERVED SYMPTOM (V5.1): rows like

      Our Barcode = <blank>   Barcode Source = Original
      Method = Barcode        Confidence = 100
      Customer Barcode = 3193504122983   Borrowed Barcode = <blank>

  i.e. a 100%-confidence "Barcode" match against a source row that has NO
  barcode at all. Manual checking confirmed these actually came from the
  material_name BRIDGE, not from the LULU catalogue.

  ROOT CAUSE — three bugs stacked:

    (a) NaN IS TRUTHY. src_barcode_display came straight off the source
        frame, so a missing barcode was float('nan'), not None. Both
        `bool(float('nan'))` and `bool('nan')` are True in Python. So the
        guard `if not src_barcode_display[src_i]:` never fired for those rows.

    (b) THE LEARNED FLAG LIVED INSIDE THAT GUARD. Because (a) made the guard
        False, `src_barcode_is_learned[src_i] = True` never ran — so FIX 6's
        `and src_barcode_is_learned[si]` test failed and 'Borrowed Barcode'
        stayed blank, while 'Barcode Source' printed 'Original' (since
        `'Original' if bc else 'N/A'` also saw truthy NaN).

    (c) BRIDGE PROVENANCE WAS LOST. _maybe_backfill() writes the customer
        barcode into barcode_to_src unconditionally. The NEXT customer with
        that barcode therefore hits the plain direct-lookup loop first and is
        labelled Method='Barcode', Confidence=100 — the fact that the barcode
        only exists because of an earlier bridge/fuzzy learn is erased.

    Note also that 'Confidence Score' in the barcode branch is HARDCODED
    (`97 if is_bridged else 100`) and 'Brand Match' / 'Weight Match' are
    computed from the description strings — none of them are evidence that a
    real catalogue barcode was involved. That is why these rows looked
    "exact" on manual inspection.

CHANGES IN THIS VERSION (search for "★ FIX 7" / "★ FIX 8" / "★ FIX 9"):

  ★ FIX 7 — NaN-proof barcode state. main() now builds a one-time snapshot
    `src_barcode_original_norm[i]` = normalize_barcode() of whatever the LULU
    source file actually contained (None if unusable), and rewrites
    src_barcode_display so NaN / 'nan' / '0' become a real None. Every
    truthiness test on a barcode is replaced by an explicit
    `normalize_barcode(...) is not None` test. The learned flag is also moved
    OUT of the display guard, so a row can be flagged learned even when its
    display value was already populated.

  ★ FIX 8 — Borrowed Barcode derived from the ORIGINAL snapshot, not from a
    mutable flag. The order-independent definition of "borrowed" is simply:
    the barcode this row matched on is NOT this source row's own original
    catalogue barcode. finalize_barcode_fields() applies that test to EVERY
    matched row (fresh or cached) and writes the customer's barcode into
    'Borrowed Barcode'. Rows bridged purely by material_name (customer has no
    barcode of its own) carry a '_bridge_bc' so the barcode that actually
    resolved them is shown instead.

  ★ FIX 9 — Bridge provenance restored. Any barcode/bridged row whose match
    was NOT against an original catalogue barcode is relabelled at finalize
    time to Method='Barcode (Bridged)', Match Status='Matched (Barcode-Bridged)',
    Confidence=97 — so it can no longer masquerade as a clean 100. A new
    'Barcode Link' column records HOW that barcode entered the index
    ('Catalogue', or the teaching method: 'Barcode (Bridged)' / 'Fuzzy' /
    'Semantic'), reconstructed from the Learned Barcodes Audit trail (which is
    cached per-customer, so it survives cache hits).

  Also: finalize no longer POPS '_src_i'. In V5.1 all_matched and
  per_file[label] hold the SAME dict objects, so the first finalize pass
  removed the key and the second pass silently skipped every row. The helper
  columns are now dropped at DataFrame-build time instead.

CAUTION ON THE BRIDGE ITSELF (unchanged behaviour, now merely visible):
  name_key_to_src is built from ALL rows in the file, keyed on
  normalize_name_key() — uppercase + whitespace collapse only. There is no
  weight or brand gate on the bridge path: `is_bridge` short-circuits past the
  FIX 1 sanity checks. Two rows sharing an identical description string but
  differing in size will bridge to the same source row. Review the
  'Learned Barcodes Audit' sheet before trusting matches built on top of
  bridged learns.

Everything else (candidate selection, fuzzy scorer, semantic layer, barcode
pre-pass, brand pre-gate, filter logic, parquet cache, embedding cache) is
UNCHANGED.

FOLDER LAYOUT (created automatically if missing)
────────────────────────────────────────────────
  C:\\Users\\HP\\Desktop\\zero1\\            <- project root (venv lives here)
  C:\\Users\\HP\\Desktop\\zero1\\data\\      <- input files + parquet cache
  C:\\Users\\HP\\Desktop\\zero1\\cache\\     <- embedding + match-result cache
  C:\\Users\\HP\\Desktop\\zero1\\output\\    <- final Excel output
"""

# ===========================================================================
# CONFIG
# ===========================================================================

# ── ★ PROJECT FOLDERS ──────────────────────────────────────────────────────
PROJECT_ROOT = r"C:\Users\HP\Desktop\zero1"
DATA_DIR     = PROJECT_ROOT + r"\data"
CACHE_DIR    = PROJECT_ROOT + r"\cache"
OUTPUT_DIR   = PROJECT_ROOT + r"\output"

# Point this at wherever your source Excel currently lives. It does NOT
# need to already be inside DATA_DIR — the parquet cache will be written
# into DATA_DIR regardless of where the original file sits.
INPUT_FILE   = r"C:\Users\HP\Downloads\categoriesandfinal_brand_standardized (1).xlsx"
INPUT_SHEET  = 0  # set to None if reading a plain CSV
OUTPUT_FILE  = OUTPUT_DIR + r"\taupdatedcate.xlsx"

# ★ CACHE — set True to ignore every cache (parquet / embeddings / match
# results) and redo the entire pipeline from scratch. Use this after you
# change the matching LOGIC itself (not just the data), since the cache
# keys don't know about code changes.
#
# NOTE: FIX 7 normalises src_barcode_display (nan -> None), which changes
# src_sig, which changes config_sig — so every per-customer match cache is
# invalidated automatically on the first run of this version. You do not
# strictly need FORCE_RECOMPUTE, but it is harmless and unambiguous.
FORCE_RECOMPUTE = False

SOURCE_CUSTOMER_LABEL = "LULU"   # rows with this source_customer are the catalogue
NAME_COL               = "material_name"          # matched on BOTH sides
BRAND_COL               = "brand"      # preferred brand field
CATEGORY_COL            = "lulu_category"                # -> hierarchy bonus (src_h1)

# ★ BARCODE — column holding the barcode/EAN/UPC on BOTH source and customer
# rows in the combined file. Set to None to disable barcode-first matching
# and fall back to pure fuzzy/semantic behaviour (identical to V3).
BARCODE_COL             = "barcode"

# ★ MATERIAL CODE — pass-through identifier column, carried into the output
# for both source and customer rows (not used in scoring/matching logic).
# Set to None to disable.
MATERIAL_CODE_COL       = "material_code"

# ★ BACKFILL — minimum confidence for a Fuzzy/Semantic match to have the
# customer's barcode "learned" against the matched LULU row. Confidence
# alone is not sufficient — see FIX 1 (weight/brand sanity gate) below.
BACKFILL_MIN_CONFIDENCE = 89

MIN_CONFIDENCE     = 75
ENSEMBLE_FUZZY     = True
USE_SEMANTIC       = True
SEMANTIC_MODEL     = "all-mpnet-base-v2"
SEMANTIC_BATCH     = 128
BRAND_FUZZY_THRESH = 75
WEIGHT_TOL         = 0.02
SAMPLE_SIZE        = None

# ★ BRAND FILTER — controls how strictly brand agreement is enforced during
# the fuzzy/semantic candidate selection (does NOT affect the barcode pass,
# which is definitive on its own regardless of brand).
REQUIRE_BRAND_MATCH       = True
ALLOW_MATCH_WITHOUT_BRAND = True

# ★ DEDUPE — drop duplicate customer rows that share the same material
# description (keeps the first occurrence of each). Mirrors the dedupe
# already applied to the LULU source side.
DEDUPE_CUSTOMER_BY_DESC = True

# ★ FIX 9 — confidence written onto a barcode row that turns out to have
# matched against a NON-original (learned/bridged) barcode. Kept distinct
# from 100 so these rows are filterable in Excel.
BRIDGED_CONFIDENCE = 97

# ===========================================================================

import re, sys, time, warnings, hashlib, json
import os as _os
import numpy as np
import pandas as pd
from collections import defaultdict
from pathlib import Path
from rapidfuzz import fuzz, process

warnings.filterwarnings("ignore")

# ── ★ CACHE / FOLDER SETUP ──────────────────────────────────────────────────
for _d in (PROJECT_ROOT, DATA_DIR, CACHE_DIR, OUTPUT_DIR):
    _os.makedirs(_d, exist_ok=True)

# helper columns carried on matched rows for post-processing only — never
# written to Excel (see write_all_results / _rows_to_df)
HELPER_COLS = ('_src_i', '_bridge_bc')


def _hash_list(items) -> str:
    """Stable short hash of a list of values (order-sensitive)."""
    h = hashlib.sha256()
    for it in items:
        h.update(str(it).encode('utf-8', errors='ignore'))
        h.update(b'\x00')
    return h.hexdigest()[:20]


def _hash_file(path) -> str:
    """Hash a file's size+mtime (fast) rather than its full contents."""
    st = _os.stat(path)
    key = f"{path}|{st.st_size}|{int(st.st_mtime)}"
    return hashlib.sha256(key.encode('utf-8')).hexdigest()[:20]


def _json_safe(obj):
    """Recursively convert numpy/pandas scalar types to plain python so json.dump doesn't choke."""
    if isinstance(obj, dict):
        return {k: _json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [_json_safe(v) for v in obj]
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating,)):
        return None if np.isnan(obj) else float(obj)
    if isinstance(obj, np.bool_):
        return bool(obj)
    if isinstance(obj, float) and np.isnan(obj):
        return None
    try:
        if pd.isna(obj):
            return None
    except (TypeError, ValueError):
        pass
    return obj


def _rows_to_df(rows) -> pd.DataFrame:
    """DataFrame from matched/unmatched rows, minus internal helper columns."""
    df = pd.DataFrame(rows)
    drop = [c for c in HELPER_COLS if c in df.columns]
    return df.drop(columns=drop) if drop else df


# ===========================================================================
# ★ PARQUET — convert input Excel to Parquet once, reuse on later runs
# ===========================================================================

def load_input_dataframe() -> pd.DataFrame:
    """
    Loads the combined input file, using a Parquet cache under DATA_DIR so
    repeated runs don't have to re-parse Excel (which is slow for large
    files). The Parquet cache is keyed off the source file's size+mtime, so
    editing/replacing the source Excel automatically triggers a fresh
    conversion.
    """
    src_path = Path(INPUT_FILE)
    if not src_path.exists():
        print(f"ERROR: input file not found: {INPUT_FILE}")
        sys.exit(1)

    file_key     = _hash_file(str(src_path))
    parquet_path = _os.path.join(DATA_DIR, f"input_{file_key}.parquet")

    if not FORCE_RECOMPUTE and _os.path.exists(parquet_path):
        try:
            print(f"  ✓ Parquet cache hit — loading → {parquet_path}")
            return pd.read_parquet(parquet_path)
        except Exception as e:
            print(f"  ⚠ Parquet cache unreadable ({e}) — reconverting from Excel")

    print("  Converting input Excel → Parquet (one-time cost) …")
    if INPUT_SHEET is not None:
        df = pd.read_excel(str(src_path), sheet_name=INPUT_SHEET, dtype={BARCODE_COL: str})
    elif str(src_path).lower().endswith('.xlsx'):
        df = pd.read_excel(str(src_path), dtype={BARCODE_COL: str})
    else:
        df = pd.read_csv(str(src_path), low_memory=False, dtype={BARCODE_COL: str})

    try:
        df.to_parquet(parquet_path, index=False)
        print(f"  ✓ Cached parquet → {parquet_path}")
    except Exception as e:
        print(f"  ⚠ Could not write parquet cache ({e}) — continuing without it")

    return df


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
    r'\bEVDAY\b': 'EVERYDAY', r'\bHNY\b': 'HONEY', r'\bBISC\b': 'BISCUIT',
    r'\bGLUC\b': 'GLUCOSE', r'\bPNAPL\b': 'PINEAPPLE', r'\bPNPL\b': 'PINEAPPLE',
    r'\bMNGO\b': 'MANGO', r'\bAPL\b': 'APPLE',
    r'\bBRZ\b': 'BRAZIL', r'\bBRZL\b': 'BRAZIL',
    r'\b&\b': 'and',
    r'\bBISCUITS\b': 'BISCUIT',
    r'\bGLD\b':  'GOLD',
    r'\bGLDN\b': 'GOLDEN',
    r'\bLNG\b':  'LONG',
    r'\bSML\b':  'SMALL',
    r'\bMED\b':  'MEDIUM',
    r'\bLRG\b':  'LARGE',
    r'\bXL\b':   'EXTRALARGE',
    r'\bORIG\b': 'ORIGINAL',
    r'\bORGNL\b':'ORIGINAL',
    r'\bTRD\b':  'TRADITIONAL',
    r'\bSTD\b':  'STANDARD',
    r'\bPRM\b':  'PREMIUM',
    r'\bPRMM\b': 'PREMIUM',
    r'\bSPC\b':  'SPECIAL',
    r'\bSPCL\b': 'SPECIAL',
    r'\bCRSPY\b':'CRISPY',
    r'\bCLSSC\b':'CLASSIC',
    r'\bCLS\b':  'CLASSIC',
    r'\bFRS\b':  'FRIES',
    r'\bFFRS\b': 'FRENCH FRIES',
    r'\bPTTO?\b':'POTATO',
    r'\bWDGS\b': 'WEDGES',
    r'\bHTDG\b': 'HOTDOG',
    r'\bNGT\b':  'NUGGET',
    r'\bNGTS\b': 'NUGGETS',
    r'\bSTC\b':  'STICKS',
    r'\bSTCKS\b':'STICKS',
    r'\bJC\b':   'JUICE',
    r'\bCNC\b':  'CONCENTRATE',
    r'\bCRBD\b': 'CARBONATED',
    r'\bMNRL\b': 'MINERAL',
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
    'LULUDAILY': 'LULUDAILY', 'ENG': 'ENGLISH', 'AHMD': 'AHMAD',
    'AHMED': 'AHMAD', 'MC CAIN': 'MCCAIN', 'MC': 'MCCAIN',
    'MCCAINS': 'MCCAINS', 'CG': 'CALIFORNIA GARDEN',
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
    'NONE', 'NAN',
    'CHOCOLATE', 'CHOCO', 'VANILLA', 'STRAWBERRY', 'RASPBERRY',
    'BLUEBERRY', 'ORANGE', 'LEMON', 'MANGO', 'PINEAPPLE', 'BANANA',
    'APPLE', 'PEACH', 'CHERRY', 'MINT', 'CARAMEL', 'COCOA', 'HONEY',
    'PLAIN', 'SPICY', 'SWEET', 'SOUR', 'SALT', 'PEPPER',
    'FLOUR', 'SUGAR', 'MILK', 'CREAM', 'YOGURT', 'BUTTER', 'GHEE',
    'BREAD', 'CAKE', 'COOKIES', 'BISCUIT', 'BISCUITS', 'WAFER',
    'CRISPY', 'CRUNCHY', 'DELUXE', 'CLASSIC', 'STANDARD', 'REGULAR',
    'FULL', 'HALF', 'QUARTER', 'WS',
    # ★ generic multi-brand prefixes (many unrelated Gulf/Arabic brands
    # share these — must never be treated as a brand signal on their own)
    'AL', 'EL', 'BIN', 'IBN', 'ABU', 'UM', 'AA',
}

HIER_STOPWORDS = {
    'AND', 'OR', 'THE', 'OF', 'IN', 'FOR', 'WITH', 'A', 'AN',
    'FRESH', 'FROZEN', 'CHILLED', 'CANNED', 'DRIED',
    'WHOLE', 'HALF', 'SLICED', 'MIXED', 'ASSORTED',
    'PRODUCT', 'PRODUCTS', 'ITEM', 'OTHER', 'MISC',
}
COLOR_TYPE_GROUPS = [
    {'WHITE', 'BROWN'},
    {'FULL', 'SKIMMED', 'SEMI'},
    {'SALTED', 'UNSALTED'},
    {'SWEETENED', 'UNSWEETENED'},
    {'REGULAR', 'DIET', 'ZERO'},
    {'LARGE', 'MEDIUM', 'SMALL'},  # egg sizes
]

PACK_RE = re.compile(
    r'\b(\d+\s*[Xx]\s*\d+(?:\s*[Xx]\s*\d+(?:\.\d+)?)?'
    r'(?:\s*(?:KG|KGS|G|GR|GM|GMS|ML|L|LT|LTR|LTRS|MG|S))?)\b',
    re.IGNORECASE,
)
SIZE_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*'
    r'(KG|KGS|G|GR|GM|GMS|GRM|GRMS|GRAMS?|MG|'
    r'L|LT|LTR|LTRS|LITRE|LITER|ML|CL|CC|S)\b',
    re.IGNORECASE,
)
GLUED_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*(KG|KGS|G|GR|GM|GMS|ML|L|LT|LTR|MG|S)\s*X\s*(\d+)\b',
    re.IGNORECASE,
)
UNIT_TABLE = {
    'KG': (1000, 'mass'), 'KGS': (1000, 'mass'),
    'G': (1, 'mass'),  'GR': (1, 'mass'),  'GM': (1, 'mass'),
    'GMS': (1, 'mass'), 'GRM': (1, 'mass'), 'GRMS': (1, 'mass'),
    'GRAM': (1, 'mass'), 'GRAMS': (1, 'mass'), 'MG': (0.001, 'mass'),
    'L': (1000, 'vol'), 'LT': (1000, 'vol'), 'LTR': (1000, 'vol'),
    'LTRS': (1000, 'vol'), 'LITRE': (1000, 'vol'), 'LITER': (1000, 'vol'),
    'ML': (1, 'vol'), 'CL': (10, 'vol'), 'CC': (1, 'vol'), 'S': (1, 'count'),
}
TOKEN_RE = re.compile(r"[A-Z0-9&]+")


# ===========================================================================
# PARSING
# ===========================================================================

CODE_RE = re.compile(r'\(\s*\d{5,}\s*\)', re.IGNORECASE)

def strip_product_codes(text: str) -> str:
    return CODE_RE.sub(' ', text)


def color_type_penalty(name_a: str, name_b: str) -> float:
    tokens_a = set(TOKEN_RE.findall(name_a.upper()))
    tokens_b = set(TOKEN_RE.findall(name_b.upper()))
    for group in COLOR_TYPE_GROUPS:
        hits_a = tokens_a & group
        hits_b = tokens_b & group
        if hits_a and hits_b and hits_a != hits_b:
            return 30.0
    return 0.0


def parse_description(text: str):
    if not isinstance(text, str):
        return ('', None, None, '')
    s = text.upper()
    s = re.sub(r'[=\.,;:]', ' ', s)
    s = re.sub(r'-', '', s)
    s = strip_product_codes(s)
    for pat, rep in ABBREV.items():
        s = re.sub(pat, rep, s)
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
    for m in GLUED_RE.finditer(s):
        unit = m.group(2).upper()
        if unit not in UNIT_TABLE:
            continue
        mult, kind = UNIT_TABLE[unit]
        try:
            sizes_found.append((float(m.group(1)) * mult, kind))
        except ValueError:
            pass
    weight_base = weight_kind = None
    if sizes_found:
        sizes_found.sort(key=lambda x: x[0])
        weight_base, weight_kind = sizes_found[0]
    pack_str = ''
    pack_tokens = list(PACK_RE.finditer(s))
    if pack_tokens:
        best_pack = max(pack_tokens, key=lambda m: len(m.group(0)))
        raw = best_pack.group(0).upper()
        raw = SIZE_RE.sub('', raw).strip()
        raw = re.sub(r'\s*[Xx]\s*', 'X', raw).strip('X')
        pack_str = raw
    else:
        gm = GLUED_RE.search(s)
        if gm:
            pack_str = gm.group(3)
    name = s
    name = GLUED_RE.sub(' ', name)
    name = PACK_RE.sub(' ', name)
    name = SIZE_RE.sub(' ', name)
    name = re.sub(r'\bX\s*\d+\b', ' ', name)
    name = re.sub(r'\bWS\b', ' ', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return (name, weight_base, weight_kind, pack_str)


# ===========================================================================
# ★ BARCODE — normalization helper
# ===========================================================================

def normalize_barcode(val):
    """
    Digits-only normalized barcode string, or None if there's effectively
    no usable barcode on this row. Handles the common junk cases:
      - NaN / None
      - '' / 'nan' / 'none' / '0' (placeholder barcodes)
      - floats read from Excel that show up as '6291234567890.0'
      - stray whitespace / hyphens inside a scanned barcode string

    ★ FIX 7 — this is now the ONLY sanctioned way to ask "does this row have
    a barcode?". Never use plain truthiness on a barcode value: NaN and the
    string 'nan' are both truthy, which is exactly what hid the V5.1 bug.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, float):
        val = f'{val:.0f}'          # 6.291e12 -> "6291123456789", .0 nahi
    s = str(val).strip()
    if s == '' or s.lower() in ('nan', 'none', '0', '0.0'):
        return None
    if 'e' in s.lower():            # "6.29e12" text form
        try:
            s = f'{float(s):.0f}'
        except ValueError:
            pass
    if s.endswith('.0'):
        s = s[:-2]
    s = re.sub(r'\D', '', s)
    s = s.lstrip("0")
    return s or None


def has_barcode(val) -> bool:
    """★ FIX 7 — explicit, NaN-proof presence test."""
    return normalize_barcode(val) is not None


def normalize_name_key(val):
    """
    Normalized key for comparing material_name across rows/customers for
    the barcode bridge: uppercase, collapsed whitespace. Deliberately NOT
    the same as parse_description() — this is a literal-text match key,
    used only to decide "is this the same material_name string as some
    other row that already has a confirmed LULU barcode".
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().upper()
    s = re.sub(r'\s+', ' ', s)
    return s if s else None


# ===========================================================================
# BRAND HELPERS
# ===========================================================================

def normalise_brand_str(raw) -> frozenset:
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
    for w in clean_name.split()[:2]:
        if w in NON_BRAND or w.isdigit() or len(w) < 3:
            continue
        return frozenset({BRAND_ABBREV.get(w, w)})
    return frozenset()


def brands_overlap(a: frozenset, b: frozenset) -> bool:
    if not a or not b:
        return False
    for x in a:
        for y in b:
            if x == y:
                return True
            if len(x) >= 5 and len(y) >= 5:
                if x in y or y in x:
                    return True
                if x[:5] == y[:5]:
                    return True
    return False


def brand_label(src_brand_raw, cust_brand_raw, threshold: int, alias_dict: dict) -> str:
    sb = str(src_brand_raw).strip().lower() if src_brand_raw else ''
    cb = str(cust_brand_raw).strip().lower() if cust_brand_raw else ''
    if not sb or not cb or sb in ('nan', 'none') or cb in ('nan', 'none'):
        return 'N/A'
    if sb == cb:
        return 'Exact'
    score = fuzz.token_sort_ratio(sb, cb)
    if score >= threshold:
        if sb not in alias_dict:
            alias_dict[sb] = {'Customer Brand': cb, 'Fuzzy Score': score}
        return 'Fuzzy'
    return 'Mismatch'


# ===========================================================================
# SIZE-BAND + WEIGHT HELPERS
# ===========================================================================

def size_band(size):
    if size is None or (isinstance(size, float) and np.isnan(size)) or size <= 0:
        return None
    return int(round(np.log10(size) * 20))


def weights_match(w1, w2, tol=WEIGHT_TOL):
    """
    Returns True (both present and within tolerance), False (both present and
    NOT within tolerance — a real conflict), or None (can't tell, since one or
    both sides have no weight at all).
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


def weight_match_label(src_w, cust_w) -> str:
    if src_w is None and cust_w is None:
        return 'N/A – no weight on either side'
    if src_w is not None and cust_w is not None:
        return (f'Matched – both {src_w:.1f}'
                if weights_match(src_w, cust_w)
                else f'Mismatch – source {src_w:.1f} vs customer {cust_w:.1f}')
    return f'Source only – {src_w:.1f}' if src_w is not None else f'Customer only – {cust_w:.1f}'


# ===========================================================================
# CATEGORY BONUS
# ===========================================================================

def category_bonus(
    cust_tokens: set,
    src_h1,
    src_h2,
    pts_per_token: float = 4.0,
    max_bonus: float = 10.0,
) -> float:
    if not cust_tokens:
        return 0.0

    hier_text = ''
    if src_h1 and not (isinstance(src_h1, float) and pd.isna(src_h1)):
        hier_text += str(src_h1).upper() + ' '
    if src_h2 and not (isinstance(src_h2, float) and pd.isna(src_h2)):
        hier_text += str(src_h2).upper()

    if not hier_text.strip():
        return 0.0

    hier_tokens = {
        tok for tok in TOKEN_RE.findall(hier_text)
        if len(tok) >= 3 and tok not in HIER_STOPWORDS
    }

    overlap = len(cust_tokens & hier_tokens)
    return min(max_bonus, overlap * pts_per_token)


# ===========================================================================
# FUZZY SCORER
# ===========================================================================

def _ensemble_score(a: str, b: str, **_) -> float:
    return (fuzz.WRatio(a, b) + fuzz.token_set_ratio(a, b) + fuzz.partial_ratio(a, b)) / 3.0


# ===========================================================================
# ★ SEMANTIC LAYER — with .npy + .json backup cache
# ===========================================================================

class SemanticMatcher:
    def __init__(self, source_names: list):
        cache_key = _hash_list(source_names) + '_' + SEMANTIC_MODEL.replace('/', '_').replace(' ', '_')
        npy_path  = _os.path.join(CACHE_DIR, f"src_emb_{cache_key}.npy")
        json_path = _os.path.join(CACHE_DIR, f"src_emb_{cache_key}.json")

        print(f"  Loading semantic model '{SEMANTIC_MODEL}' …")
        try:
            from sentence_transformers import SentenceTransformer
        except ImportError:
            raise ImportError("Run:  pip install sentence-transformers")
        self.model = SentenceTransformer(SEMANTIC_MODEL)

        self.src_emb = None

        # ── try primary cache (.npy) ────────────────────────────────────
        if not FORCE_RECOMPUTE and _os.path.exists(npy_path):
            try:
                emb = np.load(npy_path)
                if emb.shape[0] == len(source_names):
                    self.src_emb = emb
                    print(f"  ✓ Cache hit (.npy) — loaded embeddings → {npy_path}")
                else:
                    print("  ⚠ .npy cache row count mismatch — trying .json backup")
            except Exception as e:
                print(f"  ⚠ .npy cache corrupted ({e}) — trying .json backup")

        # ── fall back to .json backup if .npy failed/missing ────────────
        if self.src_emb is None and not FORCE_RECOMPUTE and _os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    raw = json.load(f)
                emb = np.array(raw, dtype=np.float32)
                if emb.shape[0] == len(source_names):
                    self.src_emb = emb
                    print(f"  ✓ Cache hit (.json backup) — loaded embeddings → {json_path}")
                    # repair the .npy so next run is fast again
                    try:
                        np.save(npy_path, self.src_emb)
                    except Exception:
                        pass
                else:
                    print("  ⚠ .json cache row count mismatch — recomputing")
            except Exception as e:
                print(f"  ⚠ .json cache corrupted ({e}) — recomputing")

        # ── both caches missing/corrupt → recompute from scratch ────────
        if self.src_emb is None:
            print(f"  Encoding {len(source_names):,} source names …")
            self.src_emb = self.model.encode(
                source_names, batch_size=SEMANTIC_BATCH,
                show_progress_bar=True, convert_to_numpy=True,
                normalize_embeddings=True,
            )
            try:
                np.save(npy_path, self.src_emb)
                print(f"  ✓ Cached embeddings (.npy) → {npy_path}")
            except Exception as e:
                print(f"  ⚠ Could not write .npy cache: {e}")
            try:
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(self.src_emb.tolist(), f)
                print(f"  ✓ Cached embeddings backup (.json) → {json_path}")
            except Exception as e:
                print(f"  ⚠ Could not write .json backup: {e}")

    def encode_batch(self, names: list) -> np.ndarray:
        return self.model.encode(
            names, batch_size=SEMANTIC_BATCH, show_progress_bar=False,
            convert_to_numpy=True, normalize_embeddings=True,
        )

    def scores_for_query(self, query_emb: np.ndarray) -> np.ndarray:
        return (self.src_emb @ query_emb) * 100.0


# ===========================================================================
# FILTER FUNCTION
# ===========================================================================

def _apply_filters(
    scores, cand_idxs,
    c_weight, c_kind_v, c_pack_v, c_brands,
    src_weight, src_kind, src_pack, src_brands,
    src_h1,
    src_h2,
    cust_tokens,
    cust_name_str,
    src_name: list,
):
    scores = scores.copy()
    for j, si in enumerate(cand_idxs):

        bonus = category_bonus(cust_tokens, src_h1[si], src_h2[si])
        scores[j] = min(100.0, scores[j] + bonus)

        penalty = color_type_penalty(cust_name_str, src_name[si])
        scores[j] = max(0.0, scores[j] - penalty)

        wm = weights_match(c_weight, src_weight[si])
        if wm is False:
            scores[j] = 0.0
            continue
        if wm is True:
            scores[j] = min(100.0, scores[j] + 5.0)
            if c_kind_v and src_kind[si] and c_kind_v != src_kind[si]:
                scores[j] = max(0.0, scores[j] - 20.0)

        scores[j] = min(100.0, scores[j] + pack_match_bonus(c_pack_v, src_pack[si]))

    order = np.argsort(scores)[::-1]
    for j in order:
        sc = float(scores[j])
        if sc < MIN_CONFIDENCE:
            break
        si = cand_idxs[j]
        if brands_overlap(c_brands, src_brands[si]):
            return si, sc, 'ok'
        # ★ BRAND FILTER — if neither side has any brand info to compare,
        # brand can't be a signal either way; fall back to name/weight score
        # alone instead of auto-rejecting a row that was never going to
        # have a brand overlap in the first place.
        if (not REQUIRE_BRAND_MATCH) or (
            ALLOW_MATCH_WITHOUT_BRAND and not c_brands and not src_brands[si]
        ):
            return si, sc, 'ok_no_brand'

    best_j  = int(np.argmax(scores))
    best_si = cand_idxs[best_j]
    best_sc = float(scores[best_j])
    return best_si, best_sc, 'brand_mismatch'


# ===========================================================================
# CUSTOMER FRAME BUILDER
# ===========================================================================

def get_customer_frame(df_all: pd.DataFrame, label: str, barcode_to_src=None) -> pd.DataFrame:
    sub = df_all[df_all['source_customer'] == label].copy()
    sub = sub.dropna(subset=[NAME_COL])

    if DEDUPE_CUSTOMER_BY_DESC:
        before = len(sub)
        if barcode_to_src and BARCODE_COL and BARCODE_COL in sub.columns:
            def _dupe_priority(bc_raw):
                bc = normalize_barcode(bc_raw)
                return 0 if (bc and bc in barcode_to_src) else 1
            sub['_dedupe_priority'] = sub[BARCODE_COL].apply(_dupe_priority)
            sub = sub.sort_values('_dedupe_priority', kind='stable')
            sub = sub.drop_duplicates(subset=[NAME_COL], keep='first')
            sub = sub.drop(columns=['_dedupe_priority'])
        else:
            sub = sub.drop_duplicates(subset=[NAME_COL])
        removed = before - len(sub)
        if removed:
            print(f"  [{label}] Dropped {removed:,} duplicate row(s) "
                  f"(same {NAME_COL}, kept the LULU-matching barcode where available)")

    sub = sub.reset_index(drop=True)
    out = pd.DataFrame({
        'desc':  sub[NAME_COL].astype(str),
        'brand': sub['_brand_resolved'],
    })
    if BARCODE_COL and BARCODE_COL in sub.columns:
        out['barcode'] = sub[BARCODE_COL].values
    else:
        out['barcode'] = None
    if MATERIAL_CODE_COL and MATERIAL_CODE_COL in sub.columns:
        out['material_code'] = sub[MATERIAL_CODE_COL].values
    else:
        out['material_code'] = None
    return out


# ===========================================================================
# MATCH ONE CUSTOMER GROUP
# ===========================================================================

def match_customer(
    label, cust,
    src_name, src_weight, src_kind, src_pack, src_brands,
    src_descraw, src_brand_display, src_buckets, all_src_idxs,
    known_src_brands, semantic,
    src_h1, src_h2,
    barcode_to_src, src_barcode_display,           # ★ BARCODE
    src_barcode_original_norm,                      # ★ FIX 7 — immutable snapshot
    src_barcode_is_learned,                         # ★ FIX 2 (now set correctly)
    src_material_code_display,                     # ★ MATERIAL CODE
    name_key_to_src, name_key_to_bc,               # ★ BRIDGE (+ FIX 8 provenance)
):
    t0 = time.time()

    print(f"\n{'='*60}")
    print(f"  Customer: {label}")
    print(f"{'='*60}")

    cust['source_file'] = label
    print(f"  Rows: {len(cust):,}")

    # ★ FIX 4 — audit trail of every barcode learned during this customer's run
    learned_audit = []

    # ★ FIX 1 — sanity-gated backfill: a Fuzzy/Semantic match can only teach a
    # barcode onto a source row if weight does not CONFLICT and brand does not
    # CONFLICT. (weights_match()/brands_overlap() being "unknown"/empty on one
    # side is fine — we only block on an actual disagreement.)
    def _maybe_backfill(cust_barcode_raw, src_i, method, conf_int,
                         cust_desc_raw=None, cust_wt=None, cust_brands_val=None):
        cust_bc_norm = normalize_barcode(cust_barcode_raw)
        if not cust_bc_norm:
            return False, None, False

        is_high_conf = (method in ('Fuzzy', 'Semantic')
                         and conf_int is not None
                         and conf_int >= BACKFILL_MIN_CONFIDENCE)
        is_bridge = method == 'Barcode (Bridged)'

        if is_high_conf:
            # weight sanity check — block on a real conflict, allow if unknown
            if weights_match(cust_wt, src_weight[src_i]) is False:
                return False, None, False
            # brand sanity check — block only if BOTH sides have brand tokens
            # and they don't overlap at all
            if cust_brands_val and src_brands[src_i] and not brands_overlap(cust_brands_val, src_brands[src_i]):
                return False, None, False

        if not (is_high_conf or is_bridge):
            return False, None, False

        newly_learned = cust_bc_norm not in barcode_to_src
        barcode_to_src[cust_bc_norm] = src_i

        # ★ FIX 7 — NaN-proof presence test, and the learned flag no longer
        # lives INSIDE the display guard. A source row that already shows some
        # barcode can still be flagged learned; and a row whose display value
        # was NaN now correctly gets both the display AND the flag.
        if not has_barcode(src_barcode_display[src_i]):
            src_barcode_display[src_i] = cust_barcode_raw
        if src_barcode_original_norm[src_i] is None:
            src_barcode_is_learned[src_i] = True

        if newly_learned:
            learned_audit.append({
                'Barcode':               cust_bc_norm,
                'Source Row Index':      src_i,
                'Our Description':       src_descraw[src_i],
                'Source Had Original':   src_barcode_original_norm[src_i] is not None,
                'Taught By Customer':    label,
                'Customer Description':  cust_desc_raw,
                'Method':                method,
                'Confidence':            conf_int,
            })

        return True, cust_barcode_raw, newly_learned

    cust['_barcode'] = cust['barcode'].apply(normalize_barcode)

    # ci -> (src_i, method_label, bridging_barcode_or_None)
    barcode_matched_cis = {}

    if barcode_to_src:
        for ci, bc in enumerate(cust['_barcode'].tolist()):
            if bc and bc in barcode_to_src:
                barcode_matched_cis[ci] = (barcode_to_src[bc], 'Barcode', None)

    n_bridged = 0
    if name_key_to_src:
        for ci in range(len(cust)):
            if ci in barcode_matched_cis:
                continue
            key = normalize_name_key(cust.at[ci, 'desc'])
            if key and key in name_key_to_src:
                # ★ FIX 8 — remember WHICH barcode resolved this name key, so a
                # row bridged purely by material_name (customer has no barcode
                # of its own) can still show what it actually matched against.
                barcode_matched_cis[ci] = (name_key_to_src[key],
                                            'Barcode (Bridged)',
                                            name_key_to_bc.get(key))
                n_bridged += 1

    n_direct = len(barcode_matched_cis) - n_bridged
    print(f"  Barcode matches: {n_direct:,} direct + {n_bridged:,} bridged "
          f"(via material_name) = {len(barcode_matched_cis):,} total  (skip fuzzy/semantic for these)")

    parsed_cust     = [parse_description(t) for t in cust['desc']]
    cust['_name']   = [p[0] for p in parsed_cust]
    cust['_weight'] = [p[1] for p in parsed_cust]
    cust['_kind']   = [p[2] for p in parsed_cust]
    cust['_pack']   = [p[3] for p in parsed_cust]
    cust['_band']   = cust['_weight'].apply(size_band)
    cust['_brands'] = [
        normalise_brand_str(b) | brand_tokens_from_desc(nm)
        for b, nm in zip(cust['brand'], cust['_name'])
    ]

    cust_name   = cust['_name'].tolist()
    cust_weight = cust['_weight'].tolist()
    cust_kind   = cust['_kind'].tolist()
    cust_pack   = cust['_pack'].tolist()
    cust_brands = cust['_brands'].tolist()

    cust_token_sets = [
        {tok for tok in TOKEN_RE.findall(nm)
         if len(tok) >= 3 and tok not in HIER_STOPWORDS}
        for nm in cust_name
    ]

    hard_unmatched_cis: set = set()
    for ci in range(len(cust)):
        if ci in barcode_matched_cis:
            continue
        cb = cust_brands[ci]
        if cb and not brands_overlap(cb, known_src_brands):
            hard_unmatched_cis.add(ci)
    print(f"  Brand pre-gate: {len(hard_unmatched_cis):,} rows → Unmatched immediately")

    cust_emb = None
    if semantic is not None:
        cust_emb = semantic.encode_batch(cust_name)

    scorer = _ensemble_score if ENSEMBLE_FUZZY else fuzz.token_set_ratio
    fuzzy_results = [None] * len(cust)

    for ci in hard_unmatched_cis:
        fuzzy_results[ci] = (None, 0.0, 'brand_not_in_catalogue', 'N/A')

    cust_groups = defaultdict(list)
    for ci in range(len(cust)):
        if ci in hard_unmatched_cis or ci in barcode_matched_cis:
            continue
        kind = cust_kind[ci]
        band = size_band(cust_weight[ci])
        kind = None if (isinstance(kind, float) and pd.isna(kind)) else kind
        cust_groups[(kind, band)].append(ci)

    total_groups = len(cust_groups)
    for g_idx, ((kind, band), cidxs) in enumerate(cust_groups.items(), 1):

        cand_idxs = []
        if band is not None:
            for d in (-1, 0, 1):
                cand_idxs += src_buckets.get((kind, band + d), [])
        cand_idxs += src_buckets.get((kind, None), [])
        cand_idxs += src_buckets.get((None, None), [])
        cand_idxs = list(dict.fromkeys(cand_idxs))

        if not cand_idxs:
            for ci in cidxs:
                fuzzy_results[ci] = (None, 0.0, 'no_candidates', 'Fuzzy')
            continue

        cand_names  = [src_name[i] for i in cand_idxs]
        query_names = [cust_name[ci] for ci in cidxs]

        all_scores = process.cdist(
            query_names, cand_names, scorer=scorer,
            workers=-1, dtype=np.float32,
        )

        for local_i, ci in enumerate(cidxs):
            row_scores = all_scores[local_i].astype(np.float64)
            best_si, best_sc, reason = _apply_filters(
                row_scores, cand_idxs,
                cust_weight[ci], cust_kind[ci], cust_pack[ci], cust_brands[ci],
                src_weight, src_kind, src_pack, src_brands,
                src_h1, src_h2,
                cust_token_sets[ci],
                cust_name[ci],
                src_name
            )
            fuzzy_results[ci] = (best_si, best_sc, reason, 'Fuzzy')

        if g_idx % 50 == 0 or g_idx == total_groups:
            print(f"  fuzzy group {g_idx}/{total_groups}  ({time.time()-t0:.1f}s)")

    def _needs_rescue(r):
        if r is None: return True
        _, sc, reason, _ = r
        return reason in ('no_candidates', 'empty_names', 'no_brand_match') \
               or (reason == 'ok' and sc < MIN_CONFIDENCE)

    needs_semantic = [
        ci for ci, r in enumerate(fuzzy_results)
        if _needs_rescue(r) and ci not in hard_unmatched_cis and ci not in barcode_matched_cis
    ]
    print(f"  Fuzzy matched: {sum(1 for r in fuzzy_results if r and r[2]=='ok' and r[1]>=MIN_CONFIDENCE):,}"
          f" | Semantic queue: {len(needs_semantic):,}")

    final_results = list(fuzzy_results)
    if semantic is not None and needs_semantic:
        for count, ci in enumerate(needs_semantic, 1):
            if cust_emb is None or not cust_name[ci]:
                continue
            best_si, best_sc, reason = _apply_filters(
                semantic.scores_for_query(cust_emb[ci]), all_src_idxs,
                cust_weight[ci], cust_kind[ci], cust_pack[ci], cust_brands[ci],
                src_weight, src_kind, src_pack, src_brands,
                src_h1, src_h2,
                cust_token_sets[ci],
                cust_name[ci],
                src_name
            )
            final_results[ci] = (best_si, best_sc, reason, 'Semantic')
            if count % 100 == 0 or count == len(needs_semantic):
                print(f"  semantic {count}/{len(needs_semantic)}  ({time.time()-t0:.1f}s)")

    matched_rows        = []
    unmatched_rows      = []
    alias_dict          = {}
    matched_src_indices = set()
    n_backfilled         = 0

    for ci in range(len(cust)):
        cust_desc_raw  = cust.at[ci, 'desc']
        cust_brand_raw = cust.at[ci, 'brand']
        cust_barcode_raw = cust.at[ci, 'barcode']
        cust_material_code_raw = cust.at[ci, 'material_code']
        cust_wt        = cust_weight[ci]
        cust_pk        = cust_pack[ci]

        if ci in barcode_matched_cis:
            src_i, bc_method, bridge_bc = barcode_matched_cis[ci]

            borrowed, borrowed_bc, newly_learned = _maybe_backfill(
                cust_barcode_raw, src_i, bc_method, None,
                cust_desc_raw=cust_desc_raw, cust_wt=cust_wt, cust_brands_val=cust_brands[ci],
            )
            if newly_learned:
                n_backfilled += 1

            src_wt = src_weight[src_i]
            src_pk = src_pack[src_i]
            src_brand_raw = src_brand_display[src_i]
            blabel = brand_label(src_brand_raw, cust_brand_raw, BRAND_FUZZY_THRESH, alias_dict)
            is_bridged = (bc_method == 'Barcode (Bridged)')

            matched_src_indices.add(src_i)
            matched_rows.append({
                # ★ FIX 3 / 8 / 9 — provisional values only; finalize_barcode_fields()
                # re-derives Our Barcode / Barcode Source / Borrowed Barcode /
                # Method / Confidence from the final state before writing.
                '_src_i':                           src_i,
                '_bridge_bc':                        bridge_bc,
                'Source File':                     label,
                'Our Brand (brand_standardized)':  src_brand_raw,
                'Our Name (parsed)':                src_name[src_i],
                'Our Material Code':                src_material_code_display[src_i],
                'Our Weight g/ml':                  src_wt,
                'Our Pack':                          src_pk,
                'Our Description (material_name)':  src_descraw[src_i],
                'Our Barcode':                       src_barcode_display[src_i],
                'Barcode Source':                    None,
                'Borrowed Barcode':                  borrowed_bc,
                'Barcode Link':                      None,
                'Customer Description':             cust_desc_raw,
                'Customer Brand':                   cust_brand_raw,
                'Customer Material Code':           cust_material_code_raw,
                'Customer Barcode':                 cust_barcode_raw,
                'Customer Name (parsed)':           cust_name[ci],
                'Customer Weight g/ml':             cust_wt,
                'Customer Pack':                     cust_pk,
                'Match Status':                     'Matched (Barcode-Bridged)' if is_bridged else 'Matched (Barcode)',
                'Confidence Score':                 BRIDGED_CONFIDENCE if is_bridged else 100,
                'Method':                           bc_method,
                'Brand Match':                      blabel,
                'Weight Match':                     weight_match_label(src_wt, cust_wt),
            })
            continue

        result = final_results[ci]

        base = {
            'Source File':          label,
            'Customer Description': cust_desc_raw,
            'Customer Brand':       cust_brand_raw,
            'Customer Material Code': cust_material_code_raw,
            'Customer Barcode':     cust_barcode_raw,
            'Customer Weight g/ml': cust_wt,
            'Customer Pack':        cust_pk,
        }

        if result is None:
            unmatched_rows.append({**base, 'Reason': 'no_result', 'Best Score': 0})
            continue

        src_i, conf, reason, method = result

        if reason in ('brand_not_in_catalogue', 'brand_mismatch') or src_i is None:
            unmatched_rows.append({**base, 'Reason': reason, 'Best Score': int(round(conf))})
            continue

        conf_int = int(round(conf))
        if conf_int < MIN_CONFIDENCE:
            unmatched_rows.append({**base, 'Reason': 'below_threshold', 'Best Score': conf_int})
            continue

        borrowed, borrowed_bc, newly_learned = _maybe_backfill(
            cust_barcode_raw, src_i, method, conf_int,
            cust_desc_raw=cust_desc_raw, cust_wt=cust_wt, cust_brands_val=cust_brands[ci],
        )
        if newly_learned:
            n_backfilled += 1

        src_wt        = src_weight[src_i]
        src_pk        = src_pack[src_i]
        src_brand_raw = src_brand_display[src_i]
        blabel        = brand_label(src_brand_raw, cust_brand_raw, BRAND_FUZZY_THRESH, alias_dict)
        status        = 'Matched (High)' if conf_int >= 85 else 'Matched (Medium)'

        matched_src_indices.add(src_i)

        matched_rows.append({
            '_src_i':                        src_i,
            '_bridge_bc':                    None,
            'Source File':                  label,
            'Our Brand (brand_standardized)': src_brand_raw,
            'Our Name (parsed)':             src_name[src_i],
            'Our Material Code':             src_material_code_display[src_i],
            'Our Weight g/ml':               src_wt,
            'Our Pack':                      src_pk,
            'Our Description (material_name)': src_descraw[src_i],
            'Our Barcode':                    src_barcode_display[src_i],
            'Barcode Source':                 None,
            'Borrowed Barcode':               borrowed_bc,
            'Barcode Link':                   None,
            'Customer Description':          cust_desc_raw,
            'Customer Brand':                cust_brand_raw,
            'Customer Material Code':        cust_material_code_raw,
            'Customer Barcode':              cust_barcode_raw,
            'Customer Name (parsed)':        cust_name[ci],
            'Customer Weight g/ml':          cust_wt,
            'Customer Pack':                 cust_pk,
            'Match Status':                  status,
            'Confidence Score':              conf_int,
            'Method':                        method,
            'Brand Match':                   blabel,
            'Weight Match':                  weight_match_label(src_wt, cust_wt),
        })

    if SAMPLE_SIZE and len(matched_rows) > SAMPLE_SIZE:
        matched_rows = matched_rows[:SAMPLE_SIZE]

    n_barcode = sum(1 for r in matched_rows if r['Method'] in ('Barcode', 'Barcode (Bridged)'))
    print(f"  → Matched: {len(matched_rows):,} (of which {n_barcode:,} via Barcode/Bridged)  "
          f"Unmatched: {len(unmatched_rows):,}  Barcodes learned: {n_backfilled:,}  ({time.time()-t0:.1f}s)")

    return matched_rows, unmatched_rows, alias_dict, matched_src_indices, learned_audit


# ===========================================================================
# ★ MATCH-RESULT CACHE — per-customer JSON, with corruption fallback
# ===========================================================================

def _customer_cache_path(label, cust, config_sig) -> str:
    cache_input = cust['desc'].astype(str).tolist() + cust['barcode'].astype(str).tolist()
    key = _hash_list(cache_input) + '_' + config_sig
    safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', str(label))[:60]
    return _os.path.join(CACHE_DIR, f"match_{safe_label}_{key}.json")


def load_customer_cache(path, barcode_to_src, src_barcode_display,
                        src_barcode_original_norm, src_barcode_is_learned):
    """Returns (m_rows, u_rows, aliases, src_idx_set, learned_audit) or None if
    cache is missing/corrupt/unusable. Also replays any barcodes this customer
    taught the catalogue last time, so downstream customers still benefit.

    ★ FIX 7 — the replay path now uses the same NaN-proof tests as
    _maybe_backfill(). In V5.1 this function set src_barcode_is_learned[si]=True
    UNCONDITIONALLY while _maybe_backfill() set it only inside a (broken) guard,
    so the same source row was labelled 'Learned' or 'Original' depending purely
    on whether that customer came from cache — an order dependency of its own.
    """
    if not _os.path.exists(path):
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            cached = json.load(f)
        m_rows        = cached['matched']
        u_rows        = cached['unmatched']
        aliases       = cached['aliases']
        src_idx_set   = set(cached['src_idx'])
        learned_audit = cached.get('learned_audit', [])
        for bc, si in cached.get('learned_barcodes', {}).items():
            barcode_to_src[bc] = si
            if not has_barcode(src_barcode_display[si]):
                src_barcode_display[si] = cached.get('learned_barcode_display', {}).get(bc)
            if src_barcode_original_norm[si] is None:
                src_barcode_is_learned[si] = True
        return m_rows, u_rows, aliases, src_idx_set, learned_audit
    except Exception as e:
        print(f"  ⚠ Match-result cache corrupted/unreadable ({e}) — recomputing this customer")
        return None


def save_customer_cache(path, m_rows, u_rows, aliases, src_idx_set, learned, learned_display, learned_audit):
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(_json_safe({
                'matched':                m_rows,
                'unmatched':               u_rows,
                'aliases':                 aliases,
                'src_idx':                 list(src_idx_set),
                'learned_barcodes':        learned,
                'learned_barcode_display': learned_display,
                'learned_audit':           learned_audit,   # ★ FIX 4
            }), f)
        print(f"  ✓ Cached match results → {path}")
    except Exception as e:
        print(f"  ⚠ Could not write match-result cache: {e}")


# ===========================================================================
# ★ FIX 3 / 8 / 9 — finalize barcode fields AFTER all customers have run
# ===========================================================================

def finalize_barcode_fields(rows, src_barcode_display, src_barcode_original_norm,
                            learned_via=None):
    """
    Re-derive every barcode-dependent output column from the FINAL state, for
    every matched row (freshly computed or loaded from cache). This is what
    makes the output independent of customer processing order and cache hits.

    The order-independent definition of a borrowed match is deliberately NOT
    "the source row's learned flag is set" (a mutable value that V5.1 failed to
    maintain). It is:

        the barcode this row matched on  !=  this source row's ORIGINAL
        catalogue barcode

    which is a comparison against an immutable snapshot taken before any
    teaching happened. That covers all three cases:
      • source row had no barcode at all  -> orig is None -> borrowed
      • source row's barcode was learned earlier by another customer -> borrowed
      • customer carries a second, different barcode for the same SKU -> borrowed

    ★ FIX 9 — a borrowed barcode match is also RELABELLED as
    'Barcode (Bridged)' with confidence BRIDGED_CONFIDENCE, so it can no longer
    sit in the output as an indistinguishable 100% 'Barcode' match. 'Barcode
    Link' records how that barcode got into the index in the first place.

    Note: does NOT pop '_src_i'. all_matched and per_file[label] hold the SAME
    dict objects, so popping made the second pass a silent no-op in V5.1. The
    helper columns are dropped at DataFrame-build time instead.
    """
    learned_via = learned_via or {}

    for row in rows:
        si = row.get('_src_i')
        if si is None:
            continue

        bc   = src_barcode_display[si]
        orig = src_barcode_original_norm[si]

        row['Our Barcode'] = bc
        row['Barcode Source'] = (
            'Original' if orig is not None
            else ('Learned (Backfilled)' if has_barcode(bc) else 'N/A')
        )

        if row.get('Method') not in ('Barcode', 'Barcode (Bridged)'):
            # Fuzzy/Semantic rows keep whatever 'Borrowed Barcode' the
            # sanity-gated backfill decided at match time.
            row.setdefault('Barcode Link', None)
            continue

        cust_bc   = normalize_barcode(row.get('Customer Barcode'))
        bridge_bc = normalize_barcode(row.get('_bridge_bc'))

        borrowed_display = None
        borrowed_key     = None

        if cust_bc is not None and cust_bc != orig:
            # customer's own barcode is what we matched on, and it is not the
            # source row's original catalogue barcode
            borrowed_display = row.get('Customer Barcode')
            borrowed_key     = cust_bc
        elif cust_bc is None and bridge_bc is not None and bridge_bc != orig:
            # bridged purely on material_name — show the barcode that resolved it
            borrowed_display = row.get('_bridge_bc')
            borrowed_key     = bridge_bc

        if borrowed_display is not None:
            row['Borrowed Barcode']  = borrowed_display
            row['Method']            = 'Barcode (Bridged)'
            row['Match Status']      = 'Matched (Barcode-Bridged)'
            row['Confidence Score']  = BRIDGED_CONFIDENCE
            row['Barcode Link']      = learned_via.get(borrowed_key, 'Learned (earlier pass)')
        else:
            # genuine catalogue barcode match
            row['Borrowed Barcode'] = None
            row['Barcode Link']     = 'Catalogue'


# ===========================================================================
# WRITE ALL RESULTS
# ===========================================================================

def write_all_results(
    all_matched, all_unmatched, alias_dict, per_file,
    total_src_rows, all_matched_src_indices, src_descraw,
    all_learned_audit,   # ★ FIX 4
):
    matched_df   = _rows_to_df(all_matched)
    unmatched_df = _rows_to_df(all_unmatched)

    src_covered_count = len(all_matched_src_indices)
    src_covered_pct   = src_covered_count / total_src_rows * 100 if total_src_rows else 0
    src_not_covered   = total_src_rows - src_covered_count
    total_cust_rows   = len(all_matched) + len(all_unmatched)

    print(f"\nWriting → {OUTPUT_FILE}")
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:

        matched_df.to_excel(writer, sheet_name='All Matched', index=False)
        if not unmatched_df.empty:
            unmatched_df.to_excel(writer, sheet_name='All Unmatched', index=False)

        for label, (m_rows, u_rows, _idx) in per_file.items():
            _rows_to_df(m_rows).to_excel(
                writer, sheet_name=f"{label} Matched"[:31], index=False)
            if u_rows:
                _rows_to_df(u_rows).to_excel(
                    writer, sheet_name=f"{label} Unmatched"[:31], index=False)

        rows = []
        rows.append({'Metric': '─── PER CUSTOMER FILE ───', 'Value': '', 'Detail': ''})
        rows.append({'Metric': '', 'Value': '', 'Detail': ''})

        for label, (m_rows, u_rows, src_idx) in per_file.items():
            t           = len(m_rows) + len(u_rows)
            match_pct   = f"{len(m_rows)/t*100:.1f}%" if t else '0%'
            unmatch_pct = f"{len(u_rows)/t*100:.1f}%" if t else '0%'
            barcode = sum(1 for r in m_rows if r['Method'] == 'Barcode')
            bridged = sum(1 for r in m_rows if r['Method'] == 'Barcode (Bridged)')
            high  = sum(1 for r in m_rows if r['Match Status'] == 'Matched (High)')
            med   = sum(1 for r in m_rows if r['Match Status'] == 'Matched (Medium)')
            fuzzy = sum(1 for r in m_rows if r['Method'] == 'Fuzzy')
            sem   = sum(1 for r in m_rows if r['Method'] == 'Semantic')
            learned_bc = sum(1 for r in m_rows if r.get('Barcode Source') == 'Learned (Backfilled)')
            borrowed_bc_n = sum(1 for r in m_rows if r.get('Borrowed Barcode'))
            cov_n = len(src_idx)
            cov_p = f"{cov_n / total_src_rows * 100:.1f}%" if total_src_rows else '0%'

            rows += [
                {'Metric': f'[{label}] Customer rows',              'Value': t,           'Detail': ''},
                {'Metric': f'[{label}] Matched',                    'Value': len(m_rows), 'Detail': match_pct},
                {'Metric': f'[{label}] Unmatched',                  'Value': len(u_rows), 'Detail': unmatch_pct},
                {'Metric': f'[{label}] via Barcode (catalogue)',    'Value': barcode,     'Detail': 'matched on an ORIGINAL LULU barcode'},
                {'Metric': f'[{label}] via Barcode (Bridged)',      'Value': bridged,     'Detail': 'matched on a learned/bridged barcode — see Borrowed Barcode'},
                {'Metric': f'[{label}] High confidence (≥85)',      'Value': high,        'Detail': f"{high/len(m_rows)*100:.1f}% of matched" if m_rows else ''},
                {'Metric': f'[{label}] Med confidence (75-84)',     'Value': med,         'Detail': f"{med/len(m_rows)*100:.1f}% of matched" if m_rows else ''},
                {'Metric': f'[{label}] via Fuzzy',                  'Value': fuzzy,       'Detail': ''},
                {'Metric': f'[{label}] via Semantic',               'Value': sem,         'Detail': ''},
                {'Metric': f'[{label}] rows shown against a Learned barcode', 'Value': learned_bc, 'Detail': 'see Learned Barcodes Audit sheet'},
                {'Metric': f'[{label}] rows with a Borrowed Barcode flag',    'Value': borrowed_bc_n, 'Detail': 'matched on a barcode that is NOT this source row\'s original'},
                {'Metric': f'[{label}] Source SKUs covered',        'Value': cov_n,       'Detail': f"{cov_p} of our {total_src_rows:,} source SKUs"},
                {'Metric': '', 'Value': '', 'Detail': ''},
            ]

        rows.append({'Metric': '─── COMBINED TOTALS (ALL FILES) ───', 'Value': '', 'Detail': ''})
        rows.append({'Metric': '', 'Value': '', 'Detail': ''})
        rows += [
            {'Metric': 'Total customer rows',    'Value': total_cust_rows,    'Detail': ''},
            {'Metric': 'Total matched',          'Value': len(all_matched),   'Detail': f"{len(all_matched)/total_cust_rows*100:.1f}% of all customer rows" if total_cust_rows else ''},
            {'Metric': 'Total unmatched',        'Value': len(all_unmatched), 'Detail': f"{len(all_unmatched)/total_cust_rows*100:.1f}% of all customer rows" if total_cust_rows else ''},
            {'Metric': 'Total barcodes learned (backfilled)', 'Value': len(all_learned_audit), 'Detail': 'see Learned Barcodes Audit sheet'},
            {'Metric': 'Total rows matched on a borrowed barcode',
             'Value': sum(1 for r in all_matched if r.get('Borrowed Barcode')),
             'Detail': 'review these — the barcode was not in the LULU catalogue'},
            {'Metric': '', 'Value': '', 'Detail': ''},
        ]

        rows.append({'Metric': '─── SOURCE CATALOGUE COVERAGE (LULU) ───', 'Value': '', 'Detail': ''})
        rows.append({'Metric': '', 'Value': '', 'Detail': ''})
        rows += [
            {'Metric': 'Total source SKUs (LULU)',      'Value': total_src_rows,    'Detail': '100%'},
            {'Metric': 'Source SKUs matched ≥1 time',   'Value': src_covered_count, 'Detail': f"{src_covered_pct:.1f}%  ← covered across all customer files"},
            {'Metric': 'Source SKUs never matched',     'Value': src_not_covered,   'Detail': f"{100 - src_covered_pct:.1f}%  ← not seen in any customer file"},
        ]

        pd.DataFrame(rows).to_excel(writer, sheet_name='Summary', index=False)

        never_idx = sorted(set(range(total_src_rows)) - all_matched_src_indices)
        if never_idx:
            pd.DataFrame({
                'Source Index':    never_idx,
                'Our Description': [src_descraw[i] for i in never_idx],
            }).to_excel(writer, sheet_name='Source Never Matched', index=False)

        pd.DataFrame([
            {'Source Brand': s, 'Customer Brand': v['Customer Brand'], 'Fuzzy Score': v['Fuzzy Score']}
            for s, v in sorted(alias_dict.items())
        ]).to_excel(writer, sheet_name='Brand Alias Dict', index=False)

        # ★ FIX 4 — Learned Barcodes Audit: every backfilled barcode, who taught
        # it, and at what confidence, so a bad learn is reviewable instead of silent.
        if all_learned_audit:
            pd.DataFrame(all_learned_audit).to_excel(
                writer, sheet_name='Learned Barcodes Audit', index=False)

    n_borrowed = sum(1 for r in all_matched if r.get('Borrowed Barcode'))
    print(f"  Total matched  : {len(matched_df):,}")
    print(f"  Total unmatched: {len(unmatched_df):,}")
    print(f"  Source coverage: {src_covered_count:,} / {total_src_rows:,} ({src_covered_pct:.1f}%)")
    print(f"  Barcodes learned (backfilled) this run: {len(all_learned_audit):,}")
    print(f"  Rows matched on a BORROWED barcode    : {n_borrowed:,}  ← now labelled 'Barcode (Bridged)'")


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    t_total = time.time()

    print(f"Project root : {PROJECT_ROOT}")
    print(f"Data dir     : {DATA_DIR}")
    print(f"Cache dir    : {CACHE_DIR}")
    print(f"Output dir   : {OUTPUT_DIR}")
    if FORCE_RECOMPUTE:
        print("  ⚠ FORCE_RECOMPUTE = True — all caches will be ignored")

    print("\nLoading combined file …")
    df = load_input_dataframe()   # ★ PARQUET — cached load

    df = df.dropna(subset=[NAME_COL]).reset_index(drop=True)
    df['source_customer'] = df['source_customer'].astype(str).str.strip()

    if BARCODE_COL in df.columns:
        print(repr(df[BARCODE_COL].dtype))
        print(df[BARCODE_COL].head(10).tolist())
        print([normalize_barcode(x) for x in df[BARCODE_COL].head(10)])

    bad_mask = df['source_customer'].str.len() > 40
    if bad_mask.any():
        print(f"  WARNING: dropping {int(bad_mask.sum())} row(s) with a corrupted "
              f"source_customer value (likely pasted SQL/junk data)")
        df = df.loc[~bad_mask].reset_index(drop=True)

    df['_brand_resolved'] = df[BRAND_COL]

    if CATEGORY_COL not in df.columns:
        print(f"  WARNING: column '{CATEGORY_COL}' not found — hierarchy bonus disabled")
        df[CATEGORY_COL] = ''

    if BARCODE_COL and BARCODE_COL not in df.columns:
        print(f"  WARNING: column '{BARCODE_COL}' not found — barcode-first matching disabled")

    src = df[df['source_customer'].str.upper() == SOURCE_CUSTOMER_LABEL.upper()].copy()
    src = src.dropna(subset=[NAME_COL]).drop_duplicates(subset=[NAME_COL]).reset_index(drop=True)
    if src.empty:
        print(f"ERROR: no rows found with source_customer == '{SOURCE_CUSTOMER_LABEL}'")
        sys.exit(1)

    parsed_src     = [parse_description(t) for t in src[NAME_COL]]
    src['_name']   = [p[0] for p in parsed_src]
    src['_weight'] = [p[1] for p in parsed_src]
    src['_kind']   = [p[2] for p in parsed_src]
    src['_pack']   = [p[3] for p in parsed_src]
    src['_band']   = src['_weight'].apply(size_band)
    src['_brands'] = [
        normalise_brand_str(b) | brand_tokens_from_desc(nm)
        for b, nm in zip(src['_brand_resolved'], src['_name'])
    ]

    known_src_brands: frozenset = frozenset(tok for bs in src['_brands'] for tok in bs)
    total_src_rows = len(src)
    print(f"  Source rows ({SOURCE_CUSTOMER_LABEL}): {total_src_rows:,}  |  Brand tokens: {len(known_src_brands):,}")

    src_buckets = defaultdict(list)
    for i, (kind, band) in enumerate(zip(src['_kind'], src['_band'])):
        kind = None if (isinstance(kind, float) and pd.isna(kind)) else kind
        src_buckets[(kind, band)].append(i)

    src_name          = src['_name'].tolist()
    src_weight        = src['_weight'].tolist()
    src_kind          = src['_kind'].tolist()
    src_pack          = src['_pack'].tolist()
    src_brands        = src['_brands'].tolist()
    src_descraw       = src[NAME_COL].tolist()
    src_brand_display = src['_brand_resolved'].tolist()
    all_src_idxs      = list(range(total_src_rows))

    src_h1 = src[CATEGORY_COL].tolist()
    src_h2 = [''] * total_src_rows
    print(f"  Hierarchy column loaded: {CATEGORY_COL}")

    if BARCODE_COL and BARCODE_COL in src.columns:
        src['_barcode'] = src[BARCODE_COL].apply(normalize_barcode)
        src_barcode_display = src[BARCODE_COL].tolist()
    else:
        src['_barcode'] = None
        src_barcode_display = [None] * total_src_rows

    # ★ FIX 7 — IMMUTABLE snapshot of what the LULU source file genuinely
    # contained, taken BEFORE any customer gets a chance to teach a barcode.
    # This is the ground truth every later "is this borrowed?" test compares
    # against, and it is never mutated.
    src_barcode_original_norm = [normalize_barcode(v) for v in src_barcode_display]

    # ★ FIX 7 — and normalise the DISPLAY list too, so a missing barcode is a
    # real None instead of a truthy float('nan') / 'nan'. This single line is
    # what stopped `if not src_barcode_display[i]:` from silently never firing.
    src_barcode_display = [v if has_barcode(v) else None for v in src_barcode_display]

    n_orig = sum(1 for x in src_barcode_original_norm if x is not None)
    print(f"  Source barcodes present in catalogue: {n_orig:,} / {total_src_rows:,} "
          f"({total_src_rows - n_orig:,} source rows have NO usable barcode)")

    # provisional flag, kept in sync but no longer authoritative — every output
    # column is re-derived from src_barcode_original_norm in finalize.
    src_barcode_is_learned = [False] * total_src_rows

    barcode_to_src = {}
    dup_barcodes = 0
    for i, bc in enumerate(src['_barcode'].tolist()):
        if not bc:
            continue
        if bc in barcode_to_src:
            dup_barcodes += 1
            continue
        barcode_to_src[bc] = i
    print(f"  Barcode index: {len(barcode_to_src):,} unique LULU barcodes"
          + (f"  ({dup_barcodes:,} duplicate barcodes in catalogue kept first occurrence)"
             if dup_barcodes else ""))

    if MATERIAL_CODE_COL and MATERIAL_CODE_COL in src.columns:
        src_material_code_display = src[MATERIAL_CODE_COL].tolist()
    else:
        if MATERIAL_CODE_COL:
            print(f"  WARNING: column '{MATERIAL_CODE_COL}' not found — 'Our Material Code' will be blank")
        src_material_code_display = [None] * total_src_rows

    # ── BRIDGE INDEX (+ ★ FIX 8 provenance) ────────────────────────────────
    name_key_to_src = {}
    name_key_to_bc  = {}   # ★ FIX 8 — which barcode resolved each name key
    if BARCODE_COL and BARCODE_COL in df.columns:
        name_to_barcodes = defaultdict(set)
        for nm, bc_raw in zip(df[NAME_COL].tolist(), df[BARCODE_COL].tolist()):
            key = normalize_name_key(nm)
            bc  = normalize_barcode(bc_raw)
            if key and bc:
                name_to_barcodes[key].add(bc)

        bridge_conflicts = 0
        for key, bcs in name_to_barcodes.items():
            resolving = sorted(bc for bc in bcs if bc in barcode_to_src)
            hits = {barcode_to_src[bc] for bc in resolving}
            if not hits:
                continue
            if len(hits) > 1:
                bridge_conflicts += 1
                continue
            name_key_to_src[key] = next(iter(hits))
            name_key_to_bc[key]  = resolving[0]

        print(f"  Barcode↔material_name bridge: {len(name_key_to_src):,} material name(s) "
              f"resolvable via a co-occurring barcode elsewhere in the file"
              + (f"  ({bridge_conflicts:,} ambiguous — multiple different LULU rows implicated, skipped)"
                 if bridge_conflicts else ""))
        print("  NOTE: the bridge has NO weight/brand gate — bridged learns bypass the "
              "FIX 1 sanity checks. Review the Learned Barcodes Audit sheet.")
    else:
        print("  Barcode↔material_name bridge: disabled (no barcode column)")

    semantic = None
    if USE_SEMANTIC:
        print("\nBuilding semantic index …")
        try:
            semantic = SemanticMatcher(src_name)   # ★ CACHE — npy+json backed
        except ImportError as e:
            print(f"  WARNING: {e}\n  Continuing with fuzzy-only.")

    customer_labels = sorted(
        df.loc[df['source_customer'].str.upper() != SOURCE_CUSTOMER_LABEL.upper(),
               'source_customer'].unique()
    )
    print(f"\nCustomer groups found: {customer_labels}")

    # ★ FIX 5 — signature of the config knobs AND the source catalogue itself,
    # so a config change OR a source-data change invalidates old per-customer
    # caches. (Because FIX 7 normalises src_barcode_display, this signature also
    # changes on the first run of V5.2 — old V5.1 caches are auto-invalidated.)
    src_sig = _hash_list(src_descraw + [str(x) for x in src_barcode_display])
    config_sig = _hash_list([
        MIN_CONFIDENCE, BRAND_FUZZY_THRESH, WEIGHT_TOL, BACKFILL_MIN_CONFIDENCE,
        REQUIRE_BRAND_MATCH, ALLOW_MATCH_WITHOUT_BRAND, DEDUPE_CUSTOMER_BY_DESC,
        SEMANTIC_MODEL, ENSEMBLE_FUZZY, USE_SEMANTIC, BRIDGED_CONFIDENCE,
        'v5.2', src_sig,
    ])

    all_matched             = []
    all_unmatched           = []
    all_aliases             = {}
    per_file                = {}
    all_matched_src_indices = set()
    all_learned_audit        = []   # ★ FIX 4

    for label in customer_labels:
        cust = get_customer_frame(df, label, barcode_to_src=barcode_to_src)
        if cust.empty:
            continue

        # ★ CACHE — check per-customer JSON cache first
        cache_path = _customer_cache_path(label, cust, config_sig)
        cached = None
        if not FORCE_RECOMPUTE:
            cached = load_customer_cache(cache_path, barcode_to_src, src_barcode_display,
                                          src_barcode_original_norm, src_barcode_is_learned)

        if cached is not None:
            print(f"\n  ✓ [{label}] Match-result cache hit — skipping recompute → {cache_path}")
            m_rows, u_rows, aliases, src_idx_set, learned_audit = cached
        else:
            barcodes_before = dict(barcode_to_src)   # snapshot to detect what THIS run learned

            m_rows, u_rows, aliases, src_idx_set, learned_audit = match_customer(
                label, cust,
                src_name, src_weight, src_kind, src_pack, src_brands,
                src_descraw, src_brand_display, src_buckets, all_src_idxs,
                known_src_brands, semantic,
                src_h1, src_h2,
                barcode_to_src, src_barcode_display,
                src_barcode_original_norm,
                src_barcode_is_learned,
                src_material_code_display,
                name_key_to_src, name_key_to_bc,
            )

            learned         = {bc: si for bc, si in barcode_to_src.items() if bc not in barcodes_before}
            learned_display = {bc: src_barcode_display[si] for bc, si in learned.items()}

            save_customer_cache(cache_path, m_rows, u_rows, aliases, src_idx_set,
                                 learned, learned_display, learned_audit)

        all_matched.extend(m_rows)
        all_unmatched.extend(u_rows)
        all_aliases.update(aliases)
        per_file[label] = (m_rows, u_rows, src_idx_set)
        all_matched_src_indices |= src_idx_set
        all_learned_audit.extend(learned_audit)

    # ★ FIX 8 / FIX 9 — now that EVERY customer (fresh or cached) has had a
    # chance to teach barcodes, re-resolve 'Our Barcode' / 'Barcode Source' /
    # 'Borrowed Barcode' / 'Method' / 'Confidence Score' / 'Barcode Link' on
    # every matched row from the final, fully-taught state. This is what
    # eliminates the stale-blank-barcode symptom AND stops a bridged/learned
    # barcode match from masquerading as a clean 100% 'Barcode' match,
    # independent of processing order or cache hit/miss.
    #
    # learned_via maps barcode -> the method that first taught it, taken from
    # the audit trail (which is cached per-customer, so it survives cache hits).
    learned_via = {}
    for a in all_learned_audit:
        learned_via.setdefault(a['Barcode'], a.get('Method') or 'Learned (unknown)')

    finalize_barcode_fields(all_matched, src_barcode_display,
                            src_barcode_original_norm, learned_via)
    # per_file rows are the SAME dict objects as in all_matched, so they are
    # already updated; this call is a harmless no-op safeguard for any row that
    # somehow only exists in one of the two collections.
    for label in per_file:
        m_rows, u_rows, src_idx_set = per_file[label]
        finalize_barcode_fields(m_rows, src_barcode_display,
                                src_barcode_original_norm, learned_via)
        per_file[label] = (m_rows, u_rows, src_idx_set)

    write_all_results(
        all_matched, all_unmatched, all_aliases, per_file,
        total_src_rows, all_matched_src_indices, src_descraw,
        all_learned_audit,
    )

    print(f"\nTotal time: {time.time()-t_total:.1f}s")
    print(f"Saved → {OUTPUT_FILE}")


if __name__ == '__main__':
    main()