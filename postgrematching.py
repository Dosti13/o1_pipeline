"""
Combined Matcher — PostgreSQL Source & Destination
====================================================
CHANGES FROM FILE-BASED VERSION
────────────────────────────────
★ Source catalogue reads from PostgreSQL table (not CSV)
★ Customer data reads from PostgreSQL table (not Excel)
★ Results write to PostgreSQL tables (not Excel)
★ Barcode match attempted FIRST — if barcode matches, skip fuzzy/semantic
★ Configurable column name for description fields (per customer table)
★ Output goes to postgres tables: matched_results, unmatched_results, summary_stats
"""

# ===========================================================================
# CONFIG — PostgreSQL CONNECTION
# ===========================================================================

PG_CONFIG = {
    'host':     '',
    'port':     5432,
    'dbname':   '',
    'user':     '',
    'password': '',
}

# ===========================================================================
# CONFIG — SOURCE TABLE (our catalogue)
# ===========================================================================

SOURCE_TABLE          = 'dim_material_master'     # Table name in postgres
SOURCE_DESC_COL       = 'material_desc'        # ← Description column name (edit here)
SOURCE_BRAND_COL      = 'mgrp_descr'           # ← Brand column name
SOURCE_BARCODE_COL    = 'barcode'              # ← Barcode column (set to None to disable)
SOURCE_PRDH1_COL      = 'prdh_descr_1'        # ← Hierarchy level 1 (set to None if absent)
SOURCE_PRDH2_COL      = 'prdh_descr_2'        # ← Hierarchy level 2 (set to None if absent)
SOURCE_SCHEMA         = 'semantic'              # Postgres schema for source table
# ===========================================================================
# CONFIG — OUTPUT TABLES
# ===========================================================================

OUTPUT_SCHEMA         = 'staging'               # Postgres schema for output tables
OUTPUT_MATCHED_TABLE  = 'grandmall_matched_results'
OUTPUT_UNMATCHED_TABLE= 'grandmall_unmatched_results'
OUTPUT_SUMMARY_TABLE  = 'grandmall_summary_stats'
OUTPUT_IF_EXISTS      = 'replace'              # 'replace' or 'append'

# ===========================================================================
# CONFIG — CUSTOMER TABLES
# ===========================================================================

CUSTOMER_TABLES = [
    {
        'table':       'stg_grand_mall',    # Postgres table name
        'label':       'Grand Mall APR 2024',  # Display label
        'desc_col':    'su_description',         # ← Description column name (edit here)
        'brand_col':   'brand',                # ← Brand column (set to None if absent)
        'barcode_col': 'barcode',              # ← Barcode column (set to None to disable)
        'schema':      'staging',               # Postgres schema
    },
    
 
    


]

# ===========================================================================
# CONFIG — MATCHING PARAMETERS
# ===========================================================================

MIN_CONFIDENCE     = 80
ENSEMBLE_FUZZY     = True
USE_SEMANTIC       = True
SEMANTIC_MODEL     = "all-mpnet-base-v2"
SEMANTIC_BATCH     = 128
BRAND_FUZZY_THRESH = 85
WEIGHT_TOL         = 0.02
SAMPLE_SIZE        = None

# ===========================================================================

import re, sys, time, warnings
import numpy as np
import pandas as pd
from collections import defaultdict
from rapidfuzz import fuzz, process
import sqlalchemy

warnings.filterwarnings("ignore")

# ── abbreviation tables (unchanged) ────────────────────────────────────────

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
    r'\b&\b': 'and', r'\bBISCUITS\b': 'BISCUIT',
    r'\bGLD\b': 'GOLD', r'\bGLDN\b': 'GOLDEN', r'\bLNG\b': 'LONG',
    r'\bSML\b': 'SMALL', r'\bMED\b': 'MEDIUM', r'\bLRG\b': 'LARGE',
    r'\bXL\b': 'EXTRALARGE', r'\bORIG\b': 'ORIGINAL',
    r'\bORGNL\b': 'ORIGINAL', r'\bTRD\b': 'TRADITIONAL',
    r'\bSTD\b': 'STANDARD', r'\bPRM\b': 'PREMIUM', r'\bPRMM\b': 'PREMIUM',
    r'\bSPC\b': 'SPECIAL', r'\bSPCL\b': 'SPECIAL',
    r'\bCRSPY\b': 'CRISPY', r'\bCLSSC\b': 'CLASSIC', r'\bCLS\b': 'CLASSIC',
    r'\bFRS\b': 'FRIES', r'\bFFRS\b': 'FRENCH FRIES',
    r'\bPTTO?\b': 'POTATO', r'\bWDGS\b': 'WEDGES', r'\bHTDG\b': 'HOTDOG',
    r'\bNGT\b': 'NUGGET', r'\bNGTS\b': 'NUGGETS',
    r'\bSTC\b': 'STICKS', r'\bSTCKS\b': 'STICKS',
    r'\bJC\b': 'JUICE', r'\bCNC\b': 'CONCENTRATE',
    r'\bCRBD\b': 'CARBONATED', r'\bMNRL\b': 'MINERAL',
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
}

HIER_STOPWORDS = {
    'AND', 'OR', 'THE', 'OF', 'IN', 'FOR', 'WITH', 'A', 'AN',
    'FRESH', 'FROZEN', 'CHILLED', 'CANNED', 'DRIED',
    'WHOLE', 'HALF', 'SLICED', 'MIXED', 'ASSORTED',
    'PRODUCT', 'PRODUCTS', 'ITEM', 'OTHER', 'MISC',
}

PACK_RE = re.compile(
    r'\b(\d+\s*[Xx]\s*\d+(?:\s*[Xx]\s*\d+(?:\.\d+)?)?'
    r'(?:\s*(?:KG|KGS|G|GR|GM|GMS|ML|L|LT|LTR|LTRS|MG))?)\b', re.IGNORECASE,
)
SIZE_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*'
    r'(KG|KGS|G|GR|GM|GMS|GRM|GRMS|GRAMS?|MG|'
    r'L|LT|LTR|LTRS|LITRE|LITER|ML|CL|CC)\b', re.IGNORECASE,
)
GLUED_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*(KG|KGS|G|GR|GM|GMS|ML|L|LT|LTR|MG)\s*X\s*(\d+)\b', re.IGNORECASE,
)
UNIT_TABLE = {
    'KG': (1000,'mass'), 'KGS': (1000,'mass'), 'G': (1,'mass'), 'GR': (1,'mass'),
    'GM': (1,'mass'), 'GMS': (1,'mass'), 'GRM': (1,'mass'), 'GRMS': (1,'mass'),
    'GRAM': (1,'mass'), 'GRAMS': (1,'mass'), 'MG': (0.001,'mass'),
    'L': (1000,'vol'), 'LT': (1000,'vol'), 'LTR': (1000,'vol'),
    'LTRS': (1000,'vol'), 'LITRE': (1000,'vol'), 'LITER': (1000,'vol'),
    'ML': (1,'vol'), 'CL': (10,'vol'), 'CC': (1,'vol'),
}
TOKEN_RE = re.compile(r"[A-Z0-9&]+")
CODE_RE  = re.compile(r'\(\s*\d{5,}\s*\)', re.IGNORECASE)


# ===========================================================================
# DATABASE HELPERS
# ===========================================================================

def make_engine():
    """Create SQLAlchemy engine from PG_CONFIG."""
    cfg = PG_CONFIG
    url = (
        f"postgresql+psycopg2://{cfg['user']}:{cfg['password']}"
        f"@{cfg['host']}:{cfg['port']}/{cfg['dbname']}"
    )
    return sqlalchemy.create_engine(url)


def load_table(engine, schema: str, table: str) -> pd.DataFrame:
    """Load a full postgres table into a DataFrame."""
    full = f'"{schema}"."{table}"'
    print(f"  Reading {full} …")
    df = pd.read_sql(f"SELECT * FROM {full}", engine)
    print(f"  → {len(df):,} rows, {len(df.columns)} columns")
    return df


def write_table(df: pd.DataFrame, engine, schema: str, table: str, if_exists='replace'):
    """Write DataFrame to a postgres table."""
    full = f"{schema}.{table}"
    df.to_sql(table, engine, schema=schema, if_exists=if_exists, index=False,
              method='multi', chunksize=500)
    print(f"  ✓ Written {len(df):,} rows → {full}")


def normalise_barcode(val) -> str | None:
    """Strip spaces/dashes, return uppercase string or None if blank."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().replace(' ', '').replace('-', '').upper()
    return s if s else None


# ===========================================================================
# PARSING  (unchanged from original)
# ===========================================================================

def strip_product_codes(text: str) -> str:
    return CODE_RE.sub(' ', text)


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
# BRAND HELPERS  (unchanged)
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
            if x == y: return True
            if len(x) >= 5 and len(y) >= 5:
                if x in y or y in x: return True
                if x[:5] == y[:5]: return True
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
# SIZE / WEIGHT / PACK HELPERS  (unchanged)
# ===========================================================================

def size_band(size):
    if size is None or (isinstance(size, float) and np.isnan(size)) or size <= 0:
        return None
    return int(round(np.log10(size) * 20))


def weights_match(w1, w2, tol=WEIGHT_TOL):
    if w1 is None or w2 is None: return None
    if (isinstance(w1, float) and np.isnan(w1)) or (isinstance(w2, float) and np.isnan(w2)): return None
    if w1 <= 0 or w2 <= 0: return None
    return (min(w1, w2) / max(w1, w2)) >= (1.0 - tol)


def pack_match_bonus(p1: str, p2: str) -> float:
    return 3.0 if (p1 and p2 and p1 == p2) else 0.0


def weight_match_label(src_w, cust_w) -> str:
    if src_w is None and cust_w is None:
        return 'N/A – no weight on either side'
    if src_w is not None and cust_w is not None:
        return (f'Matched – both {src_w:.1f}' if weights_match(src_w, cust_w)
                else f'Mismatch – source {src_w:.1f} vs customer {cust_w:.1f}')
    return f'Source only – {src_w:.1f}' if src_w is not None else f'Customer only – {cust_w:.1f}'


# ===========================================================================
# HIERARCHY INDEX  (unchanged)
# ===========================================================================

def build_hier_keyword_index(prdh1_series, prdh2_series) -> dict:
    index: dict = defaultdict(set)
    for i, (h1, h2) in enumerate(zip(prdh1_series, prdh2_series)):
        combined = (
            str(h1).upper() if (h1 and not (isinstance(h1, float) and pd.isna(h1))) else ''
        ) + ' ' + (
            str(h2).upper() if (h2 and not (isinstance(h2, float) and pd.isna(h2))) else ''
        )
        for tok in TOKEN_RE.findall(combined):
            if len(tok) <= 3 or tok in HIER_STOPWORDS:
                continue
            index[tok].add(i)
    print(f"  Hier index: {len(index):,} unique keywords")
    return dict(index)


def get_hier_candidates(cust_name_upper, hier_index, all_src_idxs, min_tokens_matched=1):
    cust_tokens = {
        tok for tok in TOKEN_RE.findall(cust_name_upper)
        if len(tok) >= 3 and tok not in HIER_STOPWORDS
    }
    matched: set = set()
    tokens_found = 0
    for tok in cust_tokens:
        if tok in hier_index:
            matched |= hier_index[tok]
            tokens_found += 1
    if tokens_found >= min_tokens_matched and matched:
        return list(matched)
    return all_src_idxs


def category_bonus(cust_tokens, src_h1, src_h2, pts_per_token=4.0, max_bonus=10.0) -> float:
    if not cust_tokens:
        return 0.0
    hier_text = ''
    if src_h1 and not (isinstance(src_h1, float) and pd.isna(src_h1)):
        hier_text += str(src_h1).upper() + ' '
    if src_h2 and not (isinstance(src_h2, float) and pd.isna(src_h2)):
        hier_text += str(src_h2).upper()
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
            show_progress_bar=True, convert_to_numpy=True, normalize_embeddings=True,
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
    src_h1, src_h2, cust_tokens,
):
    scores = scores.copy()
    for j, si in enumerate(cand_idxs):
        bonus = category_bonus(cust_tokens, src_h1[si], src_h2[si])
        scores[j] = min(100.0, scores[j] + bonus)
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
# ★ NEW — BARCODE MATCH
# ===========================================================================

def build_barcode_index(src_df: pd.DataFrame, barcode_col: str | None) -> dict:
    """
    Build dict: normalised_barcode → source row index.
    Returns empty dict if barcode_col is None or not in src_df.
    """
    if barcode_col is None or barcode_col not in src_df.columns:
        print("  Barcode index: disabled (column not configured)")
        return {}
    index = {}
    for i, val in enumerate(src_df[barcode_col]):
        bc = normalise_barcode(val)
        if bc:
            index[bc] = i          # last one wins on duplicate barcodes
    print(f"  Barcode index: {len(index):,} entries")
    return index


def try_barcode_match(
    cust_barcode_raw,
    barcode_index: dict,
) -> int | None:
    """Return source row index if barcode matches, else None."""
    if not barcode_index:
        return None
    bc = normalise_barcode(cust_barcode_raw)
    if bc is None:
        return None
    return barcode_index.get(bc)


# ===========================================================================
# MATCH ONE CUSTOMER TABLE
# ===========================================================================

def match_customer(
    cust_cfg, engine,
    src_name, src_weight, src_kind, src_pack, src_brands,
    src_descraw, src_mgrp, src_buckets, all_src_idxs,
    known_src_brands, semantic,
    hier_index, src_h1, src_h2,
    barcode_index,          # ★ NEW — barcode lookup dict
):
    label      = cust_cfg['label']
    schema     = cust_cfg.get('schema', 'public')
    table      = cust_cfg['table']
    desc_col   = cust_cfg['desc_col']       # ← configurable description column
    brand_col  = cust_cfg.get('brand_col')  # may be None
    bc_col     = cust_cfg.get('barcode_col')
    t0         = time.time()

    print(f"\n{'='*60}")
    print(f"  Customer: {label}  ({schema}.{table})")
    print(f"  Description column : {desc_col}")
    print(f"  Brand column       : {brand_col or '(none)'}")
    print(f"  Barcode column     : {bc_col or '(none)'}")
    print(f"{'='*60}")

    cust = load_table(engine, schema, table)

    # Validate required column
    if desc_col not in cust.columns:
        raise ValueError(
            f"Column '{desc_col}' not found in {schema}.{table}. "
            f"Available: {list(cust.columns)}"
        )

    cust = cust.dropna(subset=[desc_col]).drop_duplicates(subset=[desc_col]).reset_index(drop=True)
    cust['_desc']  = cust[desc_col].astype(str)
    cust['_brand'] = cust[brand_col].astype(str) if brand_col and brand_col in cust.columns else ''
    cust['_bc']    = cust[bc_col].astype(str) if bc_col and bc_col in cust.columns else None
    cust['source_file'] = label
    print(f"  Rows after dedup: {len(cust):,}")

    parsed_cust     = [parse_description(t) for t in cust['_desc']]
    cust['_name']   = [p[0] for p in parsed_cust]
    cust['_weight'] = [p[1] for p in parsed_cust]
    cust['_kind']   = [p[2] for p in parsed_cust]
    cust['_pack']   = [p[3] for p in parsed_cust]
    cust['_band']   = cust['_weight'].apply(size_band)
    cust['_brands'] = [
        normalise_brand_str(b) | brand_tokens_from_desc(nm)
        for b, nm in zip(cust['_brand'], cust['_name'])
    ]

    cust_name   = cust['_name'].tolist()
    cust_weight = cust['_weight'].tolist()
    cust_kind   = cust['_kind'].tolist()
    cust_pack   = cust['_pack'].tolist()
    cust_brands = cust['_brands'].tolist()
    cust_bc_raw = cust['_bc'].tolist() if bc_col else [None] * len(cust)

    cust_token_sets = [
        {tok for tok in TOKEN_RE.findall(nm) if len(tok) >= 3 and tok not in HIER_STOPWORDS}
        for nm in cust_name
    ]

    # Brand pre-gate
    hard_unmatched_cis: set = set()
    for ci in range(len(cust)):
        cb = cust_brands[ci]
        if cb and not brands_overlap(cb, known_src_brands):
            hard_unmatched_cis.add(ci)
    print(f"  Brand pre-gate: {len(hard_unmatched_cis):,} rows → Unmatched immediately")

    cust_emb = None
    if semantic is not None:
        cust_emb = semantic.encode_batch(cust_name)

    scorer = _ensemble_score if ENSEMBLE_FUZZY else fuzz.token_set_ratio
    fuzzy_results = [None] * len(cust)

    # Mark hard unmatched
    for ci in hard_unmatched_cis:
        fuzzy_results[ci] = (None, 0.0, 'brand_not_in_catalogue', 'N/A')

    # ★ NEW — BARCODE PASS (before fuzzy)
    barcode_matched_cis: set = set()
    if barcode_index:
        for ci in range(len(cust)):
            if ci in hard_unmatched_cis:
                continue
            src_i = try_barcode_match(cust_bc_raw[ci], barcode_index)
            if src_i is not None:
                fuzzy_results[ci] = (src_i, 100.0, 'ok', 'Barcode')
                barcode_matched_cis.add(ci)
        print(f"  Barcode matched: {len(barcode_matched_cis):,} rows → skip fuzzy/semantic")
    # ★ END NEW

    total_rows_to_match = sum(
        1 for ci in range(len(cust))
        if ci not in hard_unmatched_cis and ci not in barcode_matched_cis
    )
    processed = 0

    for ci in range(len(cust)):
        if ci in hard_unmatched_cis or ci in barcode_matched_cis:
            continue

        # Candidate selection via hierarchy index
        cand_idxs = get_hier_candidates(cust_name[ci], hier_index, all_src_idxs)

        band = size_band(cust_weight[ci])
        kind = cust_kind[ci]
        if band is not None:
            band_set = set()
            for d in (-1, 0, 1):
                band_set.update(src_buckets.get((kind, band + d), []))
            band_set.update(src_buckets.get((kind, None), []))
            band_set.update(src_buckets.get((None, None), []))
            intersection = [i for i in cand_idxs if i in band_set]
            cand_idxs = intersection if intersection else cand_idxs

        if not cand_idxs:
            fuzzy_results[ci] = (None, 0.0, 'no_candidates', 'Fuzzy')
            continue

        cand_names = [src_name[i] for i in cand_idxs]
        query_name = cust_name[ci]
        if not query_name or not cand_names:
            fuzzy_results[ci] = (None, 0.0, 'empty_names', 'Fuzzy')
            continue

        row_scores = process.cdist(
            [query_name], cand_names, scorer=scorer, workers=1, dtype=np.float32,
        )[0]

        best_si, best_sc, reason = _apply_filters(
            row_scores.astype(np.float64), cand_idxs,
            cust_weight[ci], cust_kind[ci], cust_pack[ci], cust_brands[ci],
            src_weight, src_kind, src_pack, src_brands,
            src_h1, src_h2, cust_token_sets[ci],
        )
        fuzzy_results[ci] = (best_si, best_sc, reason, 'Fuzzy')

        processed += 1
        if processed % 200 == 0 or processed == total_rows_to_match:
            print(f"  fuzzy {processed}/{total_rows_to_match}  ({time.time()-t0:.1f}s)")

    # Semantic rescue
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
                src_h1, src_h2, cust_token_sets[ci],
            )
            final_results[ci] = (best_si, best_sc, reason, 'Semantic')
            if count % 100 == 0 or count == len(needs_semantic):
                print(f"  semantic {count}/{len(needs_semantic)}  ({time.time()-t0:.1f}s)")

    # Build output rows
    matched_rows        = []
    unmatched_rows      = []
    alias_dict          = {}
    matched_src_indices = set()

    for ci in range(len(cust)):
        cust_desc_raw  = cust.at[ci, '_desc']
        cust_brand_raw = cust.at[ci, '_brand']
        cust_bc_val    = cust_bc_raw[ci] if cust_bc_raw else None
        cust_wt        = cust_weight[ci]
        cust_pk        = cust_pack[ci]
        result         = final_results[ci]

        base = {
            'source_file':          label,
            'customer_description': cust_desc_raw,
            'customer_brand':       cust_brand_raw,
            'customer_barcode':     cust_bc_val,
            'customer_weight_gml':  cust_wt,
            'customer_pack':        cust_pk,
        }

        if result is None:
            unmatched_rows.append({**base, 'reason': 'no_result', 'best_score': 0})
            continue

        src_i, conf, reason, method = result

        if reason in ('brand_not_in_catalogue', 'brand_mismatch') or src_i is None:
            unmatched_rows.append({**base, 'reason': reason, 'best_score': int(round(conf))})
            continue

        conf_int = int(round(conf))
        if conf_int < MIN_CONFIDENCE:
            unmatched_rows.append({**base, 'reason': 'below_threshold', 'best_score': conf_int})
            continue

        src_wt        = src_weight[src_i]
        src_pk        = src_pack[src_i]
        src_brand_raw = src_mgrp[src_i]
        blabel        = brand_label(src_brand_raw, cust_brand_raw, BRAND_FUZZY_THRESH, alias_dict)
        status        = 'Matched (High)' if conf_int >= 85 else 'Matched (Medium)'
        matched_src_indices.add(src_i)

        matched_rows.append({
            'source_file':            label,
            'our_brand':              src_brand_raw,
            'our_name_parsed':        src_name[src_i],
            'our_weight_gml':         src_wt,
            'our_pack':               src_pk,
            'our_description':        src_descraw[src_i],
            'materialcode':            src_mgrp[src_i],
            'customer_description':   cust_desc_raw,
            'customer_brand':         cust_brand_raw,
            'customer_barcode':       cust_bc_val,
            'customer_name_parsed':   cust_name[ci],
            'customer_weight_gml':    cust_wt,
            'customer_pack':          cust_pk,
            'match_status':           status,
            'confidence_score':       conf_int,
            'method':                 method,
            'brand_match':            blabel,
            'weight_match':           weight_match_label(src_wt, cust_wt),
        })

    if SAMPLE_SIZE and len(matched_rows) > SAMPLE_SIZE:
        matched_rows = matched_rows[:SAMPLE_SIZE]

    print(f"  → Matched: {len(matched_rows):,}  Unmatched: {len(unmatched_rows):,}  "
          f"({time.time()-t0:.1f}s)")
    return matched_rows, unmatched_rows, alias_dict, matched_src_indices


# ===========================================================================
# WRITE RESULTS TO POSTGRES
# ===========================================================================

def write_all_results(
    engine,
    all_matched, all_unmatched, alias_dict, per_file,
    total_src_rows, all_matched_src_indices, src_descraw,
):
    schema = OUTPUT_SCHEMA

    matched_df   = pd.DataFrame(all_matched)
    unmatched_df = pd.DataFrame(all_unmatched)

    src_covered_count = len(all_matched_src_indices)
    src_covered_pct   = src_covered_count / total_src_rows * 100 if total_src_rows else 0
    total_cust_rows   = len(all_matched) + len(all_unmatched)

    print(f"\nWriting results to PostgreSQL ({schema})…")

    # matched
    if not matched_df.empty:
        write_table(matched_df, engine, schema, OUTPUT_MATCHED_TABLE, OUTPUT_IF_EXISTS)

    # unmatched
    if not unmatched_df.empty:
        write_table(unmatched_df, engine, schema, OUTPUT_UNMATCHED_TABLE, OUTPUT_IF_EXISTS)

    # summary
    rows = []
    for label, (m_rows, u_rows, src_idx) in per_file.items():
        t         = len(m_rows) + len(u_rows)
        high  = sum(1 for r in m_rows if r['match_status'] == 'Matched (High)')
        med   = sum(1 for r in m_rows if r['match_status'] == 'Matched (Medium)')
        fuzzy = sum(1 for r in m_rows if r['method'] == 'Fuzzy')
        sem   = sum(1 for r in m_rows if r['method'] == 'Semantic')
        bc    = sum(1 for r in m_rows if r['method'] == 'Barcode')
        cov_n = len(src_idx)
        rows.append({
            'customer_file':       label,
            'total_customer_rows': t,
            'matched':             len(m_rows),
            'unmatched':           len(u_rows),
            'match_pct':           round(len(m_rows) / t * 100, 1) if t else 0,
            'high_confidence':     high,
            'med_confidence':      med,
            'via_barcode':         bc,
            'via_fuzzy':           fuzzy,
            'via_semantic':        sem,
            'source_skus_covered': cov_n,
            'source_coverage_pct': round(cov_n / total_src_rows * 100, 1) if total_src_rows else 0,
        })

    rows.append({
        'customer_file':       '__TOTAL__',
        'total_customer_rows': total_cust_rows,
        'matched':             len(all_matched),
        'unmatched':           len(all_unmatched),
        'match_pct':           round(len(all_matched) / total_cust_rows * 100, 1) if total_cust_rows else 0,
        'high_confidence':     None,
        'med_confidence':      None,
        'via_barcode':         None,
        'via_fuzzy':           None,
        'via_semantic':        None,
        'source_skus_covered': src_covered_count,
        'source_coverage_pct': round(src_covered_pct, 1),
    })

    write_table(pd.DataFrame(rows), engine, schema, OUTPUT_SUMMARY_TABLE, OUTPUT_IF_EXISTS)

    print(f"\n  Total matched  : {len(matched_df):,}")
    print(f"  Total unmatched: {len(unmatched_df):,}")
    print(f"  Source coverage: {src_covered_count:,} / {total_src_rows:,} ({src_covered_pct:.1f}%)")


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    t_total = time.time()

    engine = make_engine()
    print("Testing DB connection …")
    with engine.connect() as conn:
        conn.execute(sqlalchemy.text("SELECT 1"))
    print("  ✓ Connected\n")

    # ── Load source catalogue from Postgres ─────────────────────────────────
    print(f"Loading source catalogue from '{OUTPUT_SCHEMA}.{SOURCE_TABLE}' …")
    src = load_table(engine, SOURCE_SCHEMA, SOURCE_TABLE)
    src = src.dropna(subset=[SOURCE_DESC_COL]).reset_index(drop=True)

    # Hierarchy columns
    for col_cfg, col_name in [(SOURCE_PRDH1_COL, 'prdh_descr_1'), (SOURCE_PRDH2_COL, 'prdh_descr_2')]:
        if col_cfg is None or col_cfg not in src.columns:
            print(f"  WARNING: hierarchy column '{col_cfg}' not found — bonus disabled")
            src[col_name] = ''
        else:
            src[col_name] = src[col_cfg]

    parsed_src     = [parse_description(t) for t in src[SOURCE_DESC_COL]]
    src['_name']   = [p[0] for p in parsed_src]
    src['_weight'] = [p[1] for p in parsed_src]
    src['_kind']   = [p[2] for p in parsed_src]
    src['_pack']   = [p[3] for p in parsed_src]
    src['_band']   = src['_weight'].apply(size_band)
    src['_brands'] = [
        normalise_brand_str(mg) | brand_tokens_from_desc(nm)
        for mg, nm in zip(src[SOURCE_BRAND_COL], src['_name'])
    ]

    known_src_brands: frozenset = frozenset(tok for bs in src['_brands'] for tok in bs)
    total_src_rows = len(src)
    print(f"  Source rows: {total_src_rows:,}  |  Brand tokens: {len(known_src_brands):,}")

    src_buckets = defaultdict(list)
    for i, (kind, band) in enumerate(zip(src['_kind'], src['_band'])):
        kind = None if (isinstance(kind, float) and pd.isna(kind)) else kind
        src_buckets[(kind, band)].append(i)

    src_name    = src['_name'].tolist()
    src_weight  = src['_weight'].tolist()
    src_kind    = src['_kind'].tolist()
    src_pack    = src['_pack'].tolist()
    src_brands  = src['_brands'].tolist()
    src_descraw = src[SOURCE_DESC_COL].tolist()
    src_mgrp    = src[SOURCE_BRAND_COL].tolist()
    all_src_idxs = list(range(total_src_rows))

    # ★ Barcode index
    barcode_index = build_barcode_index(src, SOURCE_BARCODE_COL)

    # Hierarchy index
    print("\nBuilding hierarchy keyword index …")
    hier_index = build_hier_keyword_index(src['prdh_descr_1'], src['prdh_descr_2'])
    src_h1 = src['prdh_descr_1'].tolist()
    src_h2 = src['prdh_descr_2'].tolist()

    # Semantic index
    semantic = None
    if USE_SEMANTIC:
        print("\nBuilding semantic index …")
        try:
            semantic = SemanticMatcher(src_name)
        except ImportError as e:
            print(f"  WARNING: {e}\n  Continuing with fuzzy-only.")

    # ── Process each customer table ──────────────────────────────────────────
    all_matched             = []
    all_unmatched           = []
    all_aliases             = {}
    per_file                = {}
    all_matched_src_indices = set()

    for cust_cfg in CUSTOMER_TABLES:
        m_rows, u_rows, aliases, src_idx_set = match_customer(
            cust_cfg, engine,
            src_name, src_weight, src_kind, src_pack, src_brands,
            src_descraw, src_mgrp, src_buckets, all_src_idxs,
            known_src_brands, semantic,
            hier_index, src_h1, src_h2,
            barcode_index,
        )
        all_matched.extend(m_rows)
        all_unmatched.extend(u_rows)
        all_aliases.update(aliases)
        per_file[cust_cfg['label']] = (m_rows, u_rows, src_idx_set)
        all_matched_src_indices |= src_idx_set

    write_all_results(
        engine,
        all_matched, all_unmatched, all_aliases, per_file,
        total_src_rows, all_matched_src_indices, src_descraw,
    )

    print(f"\nTotal time: {time.time()-t_total:.1f}s")
    print(f"Done ✓")


if __name__ == '__main__':
    main()