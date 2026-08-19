"""
Combined Matcher — PostgreSQL Edition
======================================
SOURCE     : PostgreSQL table (our catalogue)
DESTINATION: PostgreSQL table (matched results written back)
CUSTOMER   : PostgreSQL table(s) (one per customer file, configurable)

MATCH STRATEGY (in order):
  1. BARCODE MATCH  — if barcode exists on both sides and matches → instant win
  2. FUZZY MATCH    — name + weight + brand + hierarchy bonus cascade
  3. SEMANTIC MATCH — rescue pass for fuzzy failures

HOW TO ADD MORE COLUMNS (read this once):
  ─────────────────────────────────────────
  Step 1 → Add the column name to the relevant CONFIG section below
            (SRC_COLUMNS, CUST_COLUMNS, or OUTPUT_COLUMNS)
  Step 2 → Pull it from the DataFrame in main() / match_customer()
            e.g.  src['_newcol'] = src['new_col_name'].tolist()
  Step 3 → Pass it into _apply_filters() or build_output_row() as needed
  Step 4 → Write it into matched_rows / unmatched_rows dicts in the
            BUILD OUTPUT ROWS section at the bottom of match_customer()
  That's it — the writer picks up any new keys automatically.
"""

# ===========================================================================
# ██████████████████████  POSTGRES CONNECTION  ████████████████████████████
# ===========================================================================

PG_SOURCE = {
    'host':     '',        # ← change
    'port':     5432,
    'dbname':   '',          # ← change
    'user':     '',        # ← change
    'password': '',    # ← change
}

# Destination can be same DB or a different one


# ===========================================================================
# ██████████████████████  SOURCE TABLE CONFIG  ████████████████████████████
# ===========================================================================
# Table that holds OUR catalogue (the "left" side of the match)

SRC_TABLE = 'dim_material_master'          # ← change to your table name
SRC_SCHEMA = 'semantic'
SRC_COLUMNS = {
    # REQUIRED — core matching columns
    'desc':    'material_desc',          # ← product description column
    'brand':   'mgrp_descr',             # ← brand column
    'barcode': 'barcode',                # ← barcode/EAN column (set None to skip)

    # OPTIONAL — hierarchy bonus columns (set None if not available)
    'prdh1':   'prdh_descr_1',           # ← category level 1
    'prdh2':   'prdh_descr_2',           # ← category level 2
    'material_code': 'material_code',     # ← any extra source column you want in output
    # ADD MORE SOURCE COLUMNS HERE
    # 'sku_code': 'sku_code',
    # 'pack_size': 'pack_size_col',
}

# ===========================================================================
# ██████████████████████  CUSTOMER TABLE CONFIG  ██████████████████████████
# ===========================================================================

CUSTOMER_TABLES = [
    {
        'label':       'Spar',          # Display label for this customer (used in output)
        'table':       'stg_spar',     # ← change
        'desc_col':    'family_text',            # ← description column in this table
        'brand_col':   'retail_article_brand_description_text',            # ← brand column in this table
        'barcode_col': None,                   # ← barcode column (set None to skip)

        # ADD MORE CUSTOMER COLUMNS HERE (they'll appear in output)
        # 'extra_cols': ['retailer_sku', 'category'],
    },
    # {
    #     'label':       'Almeera',
    #     'table':       'almeera_products',
    #     'desc_col':    'product_name',
    #     'brand_col':   'brand',
    #     'barcode_col': 'barcode',
    # },
]

# ===========================================================================
# ██████████████████████  DESTINATION TABLE CONFIG  ███████████████████████
# ===========================================================================

DEST_TABLE_MATCHED   = 'Spar_matcher_results_matched'    # ← change
DEST_TABLE_UNMATCHED = 'Spar_matcher_results_unmatched'  # ← change

# Set True to DROP and recreate tables on every run
# Set False to APPEND to existing tables
DEST_OVERWRITE = True

# ===========================================================================
# ██████████████████████  MATCH SETTINGS  █████████████████████████████████
# ===========================================================================

MIN_CONFIDENCE     = 80
ENSEMBLE_FUZZY     = True
USE_SEMANTIC       = True
SEMANTIC_MODEL     = "all-mpnet-base-v2"
SEMANTIC_BATCH     = 128
BRAND_FUZZY_THRESH = 80
WEIGHT_TOL         = 0.02
SAMPLE_SIZE        = None

# ===========================================================================
# IMPORTS
# ===========================================================================

import re, sys, time, warnings
import numpy as np
import pandas as pd
from collections import defaultdict
from rapidfuzz import fuzz, process

try:
    import psycopg2
    import psycopg2.extras
    from sqlalchemy import create_engine, text
except ImportError:
    print("ERROR: Run:  pip install psycopg2-binary sqlalchemy")
    sys.exit(1)

warnings.filterwarnings("ignore")

# ===========================================================================
# DB HELPERS
# ===========================================================================

def make_engine(cfg: dict):
    """Build a SQLAlchemy engine from a PG config dict."""
    url = (
        f"postgresql+psycopg2://{cfg['user']}:{cfg['password']}"
        f"@{cfg['host']}:{cfg['port']}/{cfg['dbname']}"
    )
    return create_engine(url, pool_pre_ping=True)


def load_table(engine, schema: str, table: str, columns: list = None) -> pd.DataFrame:
    """Load a table (or specific columns) into a DataFrame."""
    cols = ', '.join(f'"{c}"' for c in columns) if columns else '*'
    query = f'SELECT {cols} FROM "{schema}"."{table}"'
    return pd.read_sql(query, engine)


def write_df_to_pg(df: pd.DataFrame, engine, table: str, overwrite: bool):
    """Write DataFrame to PostgreSQL table."""
    if_exists = 'replace' if overwrite else 'append'
    df.to_sql(table, schema='staging', con=engine, if_exists=if_exists, index=False, method='multi', chunksize=500)
    print(f"  Written {len(df):,} rows → {table} (mode: {if_exists})")


# ===========================================================================
# ABBREVIATION DICTIONARIES  (unchanged from original)
# ===========================================================================

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
    r'\bGLD\b': 'GOLD',    r'\bGLDN\b': 'GOLDEN',
    r'\bLNG\b': 'LONG',    r'\bSML\b': 'SMALL',
    r'\bMED\b': 'MEDIUM',  r'\bLRG\b': 'LARGE',
    r'\bXL\b': 'EXTRALARGE',
    r'\bORIG\b': 'ORIGINAL', r'\bORGNL\b': 'ORIGINAL',
    r'\bTRD\b': 'TRADITIONAL', r'\bSTD\b': 'STANDARD',
    r'\bPRM\b': 'PREMIUM',  r'\bPRMM\b': 'PREMIUM',
    r'\bSPC\b': 'SPECIAL',  r'\bSPCL\b': 'SPECIAL',
    r'\bCRSPY\b': 'CRISPY', r'\bCLSSC\b': 'CLASSIC', r'\bCLS\b': 'CLASSIC',
    r'\bFRS\b': 'FRIES',    r'\bFFRS\b': 'FRENCH FRIES',
    r'\bPTTO?\b': 'POTATO', r'\bWDGS\b': 'WEDGES',
    r'\bHTDG\b': 'HOTDOG',  r'\bNGT\b': 'NUGGET', r'\bNGTS\b': 'NUGGETS',
    r'\bSTC\b': 'STICKS',   r'\bSTCKS\b': 'STICKS',
    r'\bJC\b': 'JUICE',     r'\bCNC\b': 'CONCENTRATE',
    r'\bCRBD\b': 'CARBONATED', r'\bMNRL\b': 'MINERAL',
}

BRAND_ABBREV = {
    'AME': 'AMERICANA', 'AMC': 'AMERICANA', 'AMER': 'AMERICANA',
    'AMRCNA': 'AMERICANA', 'TIFF': 'TIFFANY', 'TIFFFANY': 'TIFFANY',
    'DEEM': 'DEEMAH', 'AMT': 'AHMADTEA', 'AHMAD': 'AHMADTEA',
    'AHMADTEA': 'AHMADTEA', 'FOSTER': 'FOSTERCLARK',
    'FOSTERCLARK': 'FOSTERCLARK', 'FOSTERCLARKS': 'FOSTERCLARK',
    'GG': 'GREENGIANT', 'GREEN GIANT': 'GREENGIANT',
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
    {'LARGE', 'MEDIUM', 'SMALL'},
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
    'G': (1, 'mass'), 'GR': (1, 'mass'), 'GM': (1, 'mass'),
    'GMS': (1, 'mass'), 'GRM': (1, 'mass'), 'GRMS': (1, 'mass'),
    'GRAM': (1, 'mass'), 'GRAMS': (1, 'mass'), 'MG': (0.001, 'mass'),
    'L': (1000, 'vol'), 'LT': (1000, 'vol'), 'LTR': (1000, 'vol'),
    'LTRS': (1000, 'vol'), 'LITRE': (1000, 'vol'), 'LITER': (1000, 'vol'),
    'ML': (1, 'vol'), 'CL': (10, 'vol'), 'CC': (1, 'vol'), 'S': (1, 'count'),
}
TOKEN_RE = re.compile(r"[A-Z0-9&]+")
CODE_RE  = re.compile(r'\(\s*\d{5,}\s*\)', re.IGNORECASE)


# ===========================================================================
# BARCODE HELPERS
# ===========================================================================

def normalise_barcode(val) -> str | None:
    """Strip whitespace, leading zeros, return None if empty/null."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().lstrip('0')
    return s if s else None


def build_barcode_index(src_df: pd.DataFrame, barcode_col: str | None) -> dict:
    """
    Returns dict: normalised_barcode → source row index
    Only built if SRC_COLUMNS['barcode'] is configured.
    """
    if barcode_col is None or barcode_col not in src_df.columns:
        return {}
    index = {}
    for i, val in enumerate(src_df[barcode_col]):
        bc = normalise_barcode(val)
        if bc:
            index[bc] = i   # last one wins on duplicates
    return index


# ===========================================================================
# PARSING
# ===========================================================================

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
        try:
            sizes_found.append((float(m.group(1)) * UNIT_TABLE[unit][0], UNIT_TABLE[unit][1]))
        except ValueError:
            pass
    for m in GLUED_RE.finditer(s):
        unit = m.group(2).upper()
        if unit not in UNIT_TABLE:
            continue
        try:
            sizes_found.append((float(m.group(1)) * UNIT_TABLE[unit][0], UNIT_TABLE[unit][1]))
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

def category_bonus(cust_tokens: set, src_h1, src_h2,
                   pts_per_token: float = 4.0, max_bonus: float = 10.0) -> float:
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
    return min(max_bonus, len(cust_tokens & hier_tokens) * pts_per_token)


# ===========================================================================
# FUZZY SCORER
# ===========================================================================

def _ensemble_score(a: str, b: str, **_) -> float:
    return (fuzz.WRatio(a, b) + fuzz.token_set_ratio(a, b) + fuzz.partial_ratio(a, b)) / 3.0


# ===========================================================================
# SEMANTIC LAYER
# ===========================================================================

class SemanticMatcher:
    def __init__(self, source_names: list):
        print(f"  Loading semantic model '{SEMANTIC_MODEL}' …")
        try:
            from sentence_transformers import SentenceTransformer
        except ImportError:
            raise ImportError("Run:  pip install sentence-transformers")
        self.model = SentenceTransformer(SEMANTIC_MODEL)
        print(f"  Encoding {len(source_names):,} source names …")
        self.src_emb = self.model.encode(
            source_names, batch_size=SEMANTIC_BATCH,
            show_progress_bar=True, convert_to_numpy=True,
            normalize_embeddings=True,
        )

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
    src_h1, src_h2,
    cust_tokens,
    cust_name_str,
    src_name,
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

    best_j  = int(np.argmax(scores))
    best_si = cand_idxs[best_j]
    best_sc = float(scores[best_j])
    return best_si, best_sc, 'brand_mismatch'


# ===========================================================================
# MATCH ONE CUSTOMER TABLE
# ===========================================================================

def match_customer(
    cust_cfg,
    src_df,                 # full source DataFrame (for extra column access)
    src_name, src_weight, src_kind, src_pack, src_brands,
    src_descraw, src_mgrp,  src_buckets, all_src_idxs,
    known_src_brands, semantic,
    src_h1, src_h2,
    src_barcode_index: dict,   # normalised barcode → source row index
    src_engine,
):
    label       = cust_cfg['label']
    table       = cust_cfg['table']
    desc_col    = cust_cfg['desc_col']
    brand_col   = cust_cfg['brand_col']
    barcode_col = cust_cfg.get('barcode_col')   # may be None
    extra_cols  = cust_cfg.get('extra_cols', [])
    t0          = time.time()

    print(f"\n{'='*60}")
    print(f"  Customer: {label}  (table: {table})")
    print(f"{'='*60}")

    # Load customer table
    cols_to_load = list({desc_col, brand_col} | (
        {barcode_col} if barcode_col else set()
    ) | set(extra_cols))
    cust = load_table(src_engine, 'staging', table, cols_to_load)
    cust = cust.dropna(subset=[desc_col]).drop_duplicates(subset=[desc_col]).reset_index(drop=True)
    cust = cust.rename(columns={desc_col: 'desc', brand_col: 'brand'})
    if barcode_col and barcode_col != desc_col and barcode_col != brand_col:
        cust = cust.rename(columns={barcode_col: '_barcode'})
    elif barcode_col:
        cust['_barcode'] = cust.get('desc') if barcode_col == desc_col else cust.get('brand')
    else:
        cust['_barcode'] = None

    cust['source_file'] = label
    print(f"  Rows: {len(cust):,}")

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

    cust_name    = cust['_name'].tolist()
    cust_weight  = cust['_weight'].tolist()
    cust_kind    = cust['_kind'].tolist()
    cust_pack    = cust['_pack'].tolist()
    cust_brands  = cust['_brands'].tolist()
    cust_barcodes = [normalise_barcode(v) for v in cust['_barcode']]

    cust_token_sets = [
        {tok for tok in TOKEN_RE.findall(nm)
         if len(tok) >= 3 and tok not in HIER_STOPWORDS}
        for nm in cust_name
    ]

    # ── BARCODE PRE-MATCH ────────────────────────────────────────────────────
    # For rows where barcode hits → mark as instant match, skip fuzzy/semantic
    barcode_matched: dict = {}   # ci → source_index
    if src_barcode_index:
        for ci, bc in enumerate(cust_barcodes):
            if bc and bc in src_barcode_index:
                barcode_matched[ci] = src_barcode_index[bc]
        print(f"  Barcode matched: {len(barcode_matched):,} rows (instant)")

    # ── BRAND PRE-GATE ───────────────────────────────────────────────────────
    hard_unmatched_cis: set = set()
    for ci in range(len(cust)):
        if ci in barcode_matched:
            continue   # already matched by barcode — skip brand gate
        cb = cust_brands[ci]
        if cb and not brands_overlap(cb, known_src_brands):
            hard_unmatched_cis.add(ci)
    print(f"  Brand pre-gate: {len(hard_unmatched_cis):,} rows → Unmatched immediately")

    cust_emb = None
    if semantic is not None:
        cust_emb = semantic.encode_batch(cust_name)

    # ── FUZZY PASS ───────────────────────────────────────────────────────────
    scorer = _ensemble_score if ENSEMBLE_FUZZY else fuzz.token_set_ratio
    fuzzy_results = [None] * len(cust)

    # Barcode-matched rows get special result token (no fuzzy needed)
    for ci, si in barcode_matched.items():
        fuzzy_results[ci] = (si, 100.0, 'ok', 'Barcode')

    for ci in hard_unmatched_cis:
        fuzzy_results[ci] = (None, 0.0, 'brand_not_in_catalogue', 'N/A')

    cust_groups = defaultdict(list)
    for ci in range(len(cust)):
        if ci in barcode_matched or ci in hard_unmatched_cis:
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
                src_name,
            )
            fuzzy_results[ci] = (best_si, best_sc, reason, 'Fuzzy')

        if g_idx % 50 == 0 or g_idx == total_groups:
            print(f"  fuzzy group {g_idx}/{total_groups}  ({time.time()-t0:.1f}s)")

    # ── SEMANTIC RESCUE ───────────────────────────────────────────────────────
    def _needs_rescue(r):
        if r is None: return True
        _, sc, reason, _ = r
        return reason in ('no_candidates', 'empty_names', 'no_brand_match') \
               or (reason == 'ok' and sc < MIN_CONFIDENCE)

    needs_semantic = [
        ci for ci, r in enumerate(fuzzy_results)
        if _needs_rescue(r) and ci not in hard_unmatched_cis and ci not in barcode_matched
    ]
    print(f"  Barcode: {len(barcode_matched):,}  |  Fuzzy matched: "
          f"{sum(1 for r in fuzzy_results if r and r[2]=='ok' and r[1]>=MIN_CONFIDENCE):,}"
          f"  |  Semantic queue: {len(needs_semantic):,}")

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
                src_name,
            )
            final_results[ci] = (best_si, best_sc, reason, 'Semantic')
            if count % 100 == 0 or count == len(needs_semantic):
                print(f"  semantic {count}/{len(needs_semantic)}  ({time.time()-t0:.1f}s)")

    # ── BUILD OUTPUT ROWS ────────────────────────────────────────────────────
    matched_rows        = []
    unmatched_rows      = []
    alias_dict          = {}
    matched_src_indices = set()

    for ci in range(len(cust)):
        cust_desc_raw  = cust.at[ci, 'desc']
        cust_brand_raw = cust.at[ci, 'brand']
        cust_wt        = cust_weight[ci]
        cust_pk        = cust_pack[ci]
        cust_bc        = cust_barcodes[ci]
        result         = final_results[ci]

        # Extra columns (if configured in CUSTOMER_TABLES → extra_cols)
        extra = {col: cust.at[ci, col] for col in extra_cols if col in cust.columns}

        base = {
            'source_file':           label,
            'customer_description':  cust_desc_raw,
            'customer_brand':        cust_brand_raw,
            'customer_barcode':      cust_bc,
            'customer_weight_gml':   cust_wt,
            'customer_pack':         cust_pk,
              # any extra customer columns land here automatically
        }

        if result is None:
            unmatched_rows.append({**base, 'reason': 'no_result', 'best_score': 0, 'method': 'N/A'})
            continue

        src_i, conf, reason, method = result

        if reason in ('brand_not_in_catalogue', 'brand_mismatch') or src_i is None:
            unmatched_rows.append({**base, 'reason': reason, 'best_score': int(round(conf)), 'method': method})
            continue

        conf_int = int(round(conf))
        if conf_int < MIN_CONFIDENCE and method != 'Barcode':
            unmatched_rows.append({**base, 'reason': 'below_threshold', 'best_score': conf_int, 'method': method})
            continue

        src_wt        = src_weight[src_i]
        src_pk        = src_pack[src_i]
        src_brand_raw = src_mgrp[src_i]
        blabel        = brand_label(src_brand_raw, cust_brand_raw, BRAND_FUZZY_THRESH, alias_dict)
        status        = ('Matched (Barcode)' if method == 'Barcode'
                         else 'Matched (High)' if conf_int >= 85
                         else 'Matched (Medium)')

        matched_src_indices.add(src_i)

        matched_rows.append({
            **base,
            # Source side
            'our_brand':             src_brand_raw,
            'our_name_parsed':       src_name[src_i],
            'our_weight_gml':        src_wt,
            'our_pack':              src_pk,
            'our_description':       src_descraw[src_i],
            'materialcode':          src_df.at[src_i, 'material_code'],
        
            # Match metadata
            'match_status':          status,
            'confidence_score':      conf_int,
            'method':                method,
            'brand_match':           blabel,
            'weight_match':          weight_match_label(src_wt, cust_wt),

            # ─────────────────────────────────────────────────────────────────
            # TO ADD MORE OUTPUT COLUMNS:
            #   1. Pull the value from src_df or cust at this point
            #   2. Add a new key:value pair below
            # Example:
            #   'our_sku_code': src_df.at[src_i, 'sku_code'],
            # ─────────────────────────────────────────────────────────────────
        })

    if SAMPLE_SIZE and len(matched_rows) > SAMPLE_SIZE:
        matched_rows = matched_rows[:SAMPLE_SIZE]

    print(f"  → Matched: {len(matched_rows):,}  Unmatched: {len(unmatched_rows):,}  "
          f"({time.time()-t0:.1f}s)")

    return matched_rows, unmatched_rows, alias_dict, matched_src_indices


# ===========================================================================
# WRITE RESULTS TO POSTGRES
# ===========================================================================

def write_results_to_pg(
    all_matched, all_unmatched, dest_engine,
    total_src_rows, all_matched_src_indices,
):
    print(f"\nWriting results to PostgreSQL …")

    matched_df   = pd.DataFrame(all_matched)
    unmatched_df = pd.DataFrame(all_unmatched)

    if not matched_df.empty:
        write_df_to_pg(matched_df, dest_engine, DEST_TABLE_MATCHED, DEST_OVERWRITE)

    if not unmatched_df.empty:
        write_df_to_pg(unmatched_df, dest_engine, DEST_TABLE_UNMATCHED, DEST_OVERWRITE)

    src_covered     = len(all_matched_src_indices)
    src_covered_pct = src_covered / total_src_rows * 100 if total_src_rows else 0
    total_cust      = len(all_matched) + len(all_unmatched)

    print(f"\n{'─'*50}")
    print(f"  Total customer rows : {total_cust:,}")
    print(f"  Matched             : {len(all_matched):,}  ({len(all_matched)/total_cust*100:.1f}%)" if total_cust else "")
    print(f"  Unmatched           : {len(all_unmatched):,}")
    print(f"  Source SKU coverage : {src_covered:,} / {total_src_rows:,}  ({src_covered_pct:.1f}%)")
    print(f"  Matched table       : {DEST_TABLE_MATCHED}")
    print(f"  Unmatched table     : {DEST_TABLE_UNMATCHED}")
    print(f"{'─'*50}")


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    t_total = time.time()

    # ── DB connections ──────────────────────────────────────────────────────
    print("Connecting to source PostgreSQL …")
    src_engine  = make_engine(PG_SOURCE)
    print("Connecting to destination PostgreSQL …")
    dest_engine = make_engine(PG_SOURCE)

    # ── Load source catalogue ───────────────────────────────────────────────
    print(f"\nLoading source table: {SRC_TABLE} …")

    # Collect all column names we need from source
    src_cols_needed = list(filter(None, [
        SRC_COLUMNS['desc'],
        SRC_COLUMNS['brand'],
        SRC_COLUMNS.get('barcode'),
        SRC_COLUMNS.get('prdh1'),
        SRC_COLUMNS.get('prdh2'),
        SRC_COLUMNS.get('material_code'),
        # ADD MORE SOURCE COLUMNS HERE:
        # 'sku_code',
    ]))
    src = load_table(src_engine, SRC_SCHEMA, SRC_TABLE, src_cols_needed)
    src = src.dropna(subset=[SRC_COLUMNS['desc']]).reset_index(drop=True)

    # Rename to internal names
    src = src.rename(columns={SRC_COLUMNS['desc']: 'material_desc', SRC_COLUMNS['brand']: 'mgrp_descr'})

    # Hierarchy columns — fill '' if not configured or missing
    for key, internal in [('prdh1', 'prdh_descr_1'), ('prdh2', 'prdh_descr_2')]:
        original_col = SRC_COLUMNS.get(key)
        if original_col and original_col in src.columns:
            src = src.rename(columns={original_col: internal})
        else:
            print(f"  WARNING: prdh column '{original_col}' not found — hierarchy bonus disabled")
            src[internal] = ''

    parsed_src     = [parse_description(t) for t in src['material_desc']]
    src['_name']   = [p[0] for p in parsed_src]
    src['_weight'] = [p[1] for p in parsed_src]
    src['_kind']   = [p[2] for p in parsed_src]
    src['_pack']   = [p[3] for p in parsed_src]
    src['_band']   = src['_weight'].apply(size_band)
    src['_brands'] = [
        normalise_brand_str(mg) | brand_tokens_from_desc(nm)
        for mg, nm in zip(src['mgrp_descr'], src['_name'])
    ]

    known_src_brands: frozenset = frozenset(tok for bs in src['_brands'] for tok in bs)
    total_src_rows = len(src)
    print(f"  Source rows: {total_src_rows:,}  |  Brand tokens: {len(known_src_brands):,}")

    # Barcode index
    barcode_col_in_src = SRC_COLUMNS.get('barcode')
    if barcode_col_in_src and barcode_col_in_src in src.columns:
        src_barcode_index = build_barcode_index(src, barcode_col_in_src)
        print(f"  Barcode index built: {len(src_barcode_index):,} unique barcodes")
    else:
        src_barcode_index = {}
        print("  Barcode index: DISABLED (no barcode column configured)")

    src_buckets = defaultdict(list)
    for i, (kind, band) in enumerate(zip(src['_kind'], src['_band'])):
        kind = None if (isinstance(kind, float) and pd.isna(kind)) else kind
        src_buckets[(kind, band)].append(i)

    src_name    = src['_name'].tolist()
    src_weight  = src['_weight'].tolist()
    src_kind    = src['_kind'].tolist()
    src_pack    = src['_pack'].tolist()
    src_brands  = src['_brands'].tolist()
    src_descraw = src['material_desc'].tolist()
    src_mgrp    = src['mgrp_descr'].tolist()
    src_h1      = src['prdh_descr_1'].tolist()
    src_h2      = src['prdh_descr_2'].tolist()
    all_src_idxs = list(range(total_src_rows))

    # ── Semantic index ──────────────────────────────────────────────────────
    semantic = None
    if USE_SEMANTIC:
        print("\nBuilding semantic index …")
        try:
            semantic = SemanticMatcher(src_name)
        except ImportError as e:
            print(f"  WARNING: {e}\n  Continuing with fuzzy-only.")

    # ── Run matching per customer table ─────────────────────────────────────
    all_matched             = []
    all_unmatched           = []
    all_matched_src_indices = set()

    for cust_cfg in CUSTOMER_TABLES:
        m_rows, u_rows, _, src_idx_set = match_customer(
            cust_cfg,
            src,
            src_name, src_weight, src_kind, src_pack, src_brands,
            src_descraw, src_mgrp, src_buckets, all_src_idxs,
            known_src_brands, semantic,
            src_h1, src_h2,
            src_barcode_index,
            src_engine,
        )
        all_matched.extend(m_rows)
        all_unmatched.extend(u_rows)
        all_matched_src_indices |= src_idx_set

    # ── Write to destination ─────────────────────────────────────────────────
    write_results_to_pg(
        all_matched, all_unmatched, dest_engine,
        total_src_rows, all_matched_src_indices,
    )

    print(f"\nTotal time: {time.time()-t_total:.1f}s")


if __name__ == '__main__':
    main()