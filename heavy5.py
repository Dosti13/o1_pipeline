"""
Combined Matcher — NAME + WEIGHT + BRAND  (Fuzzy → Semantic cascade)
=====================================================================
v3 IMPROVEMENTS:
  1. Brand mismatch no longer blocks match — description similarity still used
  2. Brackets, parentheses, extra spaces removed from description before matching
  3. Leading numeric codes (part numbers) split from description cleanly:
       - Weight parsed from description if unit present
       - Numeric prefix removed from name so fuzzy/semantic match works on words
  4. Brand match is now a SCORING BONUS, not a hard gate
"""

# ===========================================================================
# CONFIG — SOURCE (never changes)
# ===========================================================================

OUR_FILE    = r"C:\Users\HP\Desktop\functionand sp.csv"
OUTPUT_FILE = r"C:\Users\HP\Desktop\c.xlsx"

MIN_CONFIDENCE     = 75
ENSEMBLE_FUZZY     = True
USE_SEMANTIC       = True
SEMANTIC_MODEL     = "all-MiniLM-L6-v2"
SEMANTIC_BATCH     = 128
BRAND_FUZZY_THRESH = 75
WEIGHT_TOL         = 0.05
SAMPLE_SIZE        = 500

# Brand bonus/penalty config
BRAND_MATCH_BONUS    = 8.0   # added to score when brand matches
BRAND_MISMATCH_PEN   = 5.0   # subtracted when brand explicitly mismatches (both sides have brand)

# ===========================================================================
# CUSTOMER FILES
# ===========================================================================

CUSTOMER_FILES = [
    {
        'path':   r"C:\Users\HP\Downloads\fwqniemtselloutdashboardsession2\C4 sep 24.xlsx",
        'label':  'C4',
        'reader': 'read_c4',
    },
 
]

# ===========================================================================

import re, sys, time, warnings
import numpy as np
import pandas as pd
from collections import defaultdict
from pathlib import Path
from rapidfuzz import fuzz, process

warnings.filterwarnings("ignore")

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
    r'\bLNG\b': 'LONG',      # so "GOLD LNG" → "GOLD LONG"
    r'\bGLD\b': 'GOLD',      # so "GLD" → "GOLD" on customer side
    r'\bMCCAIN\b': 'MC CAIN'
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
    'AHMED': 'AHMAD',
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
GLUED_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*(KG|KGS|G|GR|GM|GMS|ML|L|LT|LTR|MG)\s*X\s*(\d+)\b',
    re.IGNORECASE,
)
UNIT_TABLE = {
    'KG': (1000, 'mass'), 'KGS': (1000, 'mass'),
    'G': (1, 'mass'),  'GR': (1, 'mass'),  'GM': (1, 'mass'),
    'GMS': (1, 'mass'), 'GRM': (1, 'mass'), 'GRMS': (1, 'mass'),
    'GRAM': (1, 'mass'), 'GRAMS': (1, 'mass'), 'MG': (0.001, 'mass'),
    'L': (1000, 'vol'), 'LT': (1000, 'vol'), 'LTR': (1000, 'vol'),
    'LTRS': (1000, 'vol'), 'LITRE': (1000, 'vol'), 'LITER': (1000, 'vol'),
    'ML': (1, 'vol'), 'CL': (10, 'vol'), 'CC': (1, 'vol'),
}
TOKEN_RE = re.compile(r"[A-Z0-9&]+")

# ── NEW: patterns that identify a leading part-number / code prefix ──────────
# e.g. "30005-20081", "126.326", "99004-17001", "104.792-10", "888 ", "4H "
LEADING_CODE_RE = re.compile(
    r'^(\s*'
    r'(?:'
    r'\d[\d\.\-]+[A-Z0-9\-]*'    # starts with digit, then dots/dashes/alphanum
    r'|[A-Z]\d[A-Z0-9\-]*'       # letter then digit prefix like "4H"
    r')'
    r'\s*[-–]?\s*)',
    re.IGNORECASE,
)

# ── Bracket cleaner: removes code/label brackets, unwraps weight brackets ───
# e.g. (308203) → removed entirely   |   (400G X 2) → "400G X 2" kept
BRACKET_RE = re.compile(r'[\[\]\{\}]')   # fallback for square/curly brackets


# ===========================================================================
# DESCRIPTION PRE-CLEANER  (NEW)
# ===========================================================================

def clean_description(text: str) -> str:
    """
    Remove brackets, collapse whitespace, strip leading part-number codes.
    Returns cleaned string ready for parse_description().
    """
    if not isinstance(text, str):
        return ''
    s = text.strip()

    # 1. Smart parenthesis handling:
    #    (308203) / (A) / (3LAYER) → drop entirely  (SKU code / label, no value)
    #    (400G X 2)                → unwrap to content (weight info must survive)
    def _bracket_replacer(m):
        inner = m.group(0)[1:-1]
        if SIZE_RE.search(inner):       # contains a weight unit → unwrap
            return ' ' + inner + ' '
        return ' '                      # pure code / label → delete
    s = re.sub(r'\(([^)]{0,30})\)', _bracket_replacer, s)
    s = BRACKET_RE.sub(' ', s)         # remove any leftover [ ] { }

    # 2. Collapse multiple spaces / tabs
    s = re.sub(r'\s+', ' ', s).strip()

    # 3. Remove leading numeric/code prefix that is NOT a weight
    #    e.g. "30005-20081 THERMISTOR STEAM" → "THERMISTOR STEAM"
    #    e.g. "888 TOMATO PASTE 800X12"      → "TOMATO PASTE 800X12"
    #    BUT  "500GR FLOUR"                  → keep (has unit)
    #    Strategy: strip the leading token only if it contains no weight unit
    m = LEADING_CODE_RE.match(s)
    if m:
        prefix = m.group(1).strip()
        remainder = s[m.end():].strip()
        # Only strip if the prefix itself carries no weight unit
        if remainder and not SIZE_RE.search(prefix):
            s = remainder

    return s


# ===========================================================================
# PARSING
# ===========================================================================

def parse_description(text: str):
    """
    Returns (clean_name, weight_base, weight_kind, pack_str).
    Accepts already-cleaned or raw text.
    """
    if not isinstance(text, str):
        return ('', None, None, '')

    # Apply bracket/code cleaning first
    text = clean_description(text)

    s = text.upper()
    s = re.sub(r'[=\.,;:]', ' ', s)
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

    # NEW: remove any residual standalone number tokens (not attached to units)
    # so "888 TOMATO PASTE" → "TOMATO PASTE" after code strip already handled it,
    # but catch any leftover isolated integers
    name = re.sub(r'(?<!\w)\d{1,4}(?!\w)', ' ', name)

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
# FILTER FUNCTION  (v3: brand is bonus/penalty, NOT a hard gate)
# ===========================================================================

def _apply_filters(
    scores, cand_idxs,
    c_weight, c_kind_v, c_pack_v, c_brands,
    src_weight, src_kind, src_pack, src_brands,
):
    """
    Apply weight filtering and brand scoring adjustments.

    KEY CHANGE from v2:
      - Brand mismatch no longer zeroes out the score or blocks matching.
      - Brand overlap → +BRAND_MATCH_BONUS points.
      - Both sides have brands but no overlap → -BRAND_MISMATCH_PEN points.
      - One or both sides have no brand → no adjustment (neutral).
      - Weight mismatch is still a hard reject (physical impossibility).
    """
    scores = scores.copy()

    for j, si in enumerate(cand_idxs):
        # ── Weight gate (hard) ──────────────────────────────────────────────
        wm = weights_match(c_weight, src_weight[si])
        if wm is False:
            scores[j] = 0.0
            continue
        if wm is True:
            scores[j] = min(100.0, scores[j] + 5.0)
            if c_kind_v and src_kind[si] and c_kind_v != src_kind[si]:
                scores[j] = max(0.0, scores[j] - 20.0)

        # ── Pack bonus ──────────────────────────────────────────────────────
        scores[j] = min(100.0, scores[j] + pack_match_bonus(c_pack_v, src_pack[si]))

        # ── Brand soft adjustment (NEW) ─────────────────────────────────────
        if c_brands and src_brands[si]:
            # Both sides have brand info
            if brands_overlap(c_brands, src_brands[si]):
                scores[j] = min(100.0, scores[j] + BRAND_MATCH_BONUS)
            else:
                scores[j] = max(0.0, scores[j] - BRAND_MISMATCH_PEN)
        # If either side has no brand info → neutral (no change)

    # Pick the best candidate regardless of brand
    best_j  = int(np.argmax(scores))
    best_si = cand_idxs[best_j]
    best_sc = float(scores[best_j])

    # Determine reason label
    if best_sc < MIN_CONFIDENCE:
        reason = 'below_threshold'
    elif c_brands and src_brands[best_si] and brands_overlap(c_brands, src_brands[best_si]):
        reason = 'ok'
    elif c_brands and src_brands[best_si]:
        reason = 'brand_mismatch_but_matched'   # matched on desc only
    else:
        reason = 'ok'   # no brand info to compare

    return best_si, best_sc, reason


# ===========================================================================
# CUSTOMER FILE READERS
# ===========================================================================

def read_almeera(path):
    df = pd.read_excel(path, usecols=['Product name', 'Brand'])
    df = df.dropna(subset=['Product name']).drop_duplicates(subset=['Product name']).reset_index(drop=True)
    return df.rename(columns={'Product name': 'desc', 'Brand': 'brand'})

def read_c4(path):
    df = pd.read_excel(path, sheet_name='Items', usecols=['Product name', 'Brand'])
    df = df.dropna(subset=['Product name']).drop_duplicates(subset=['Product name']).reset_index(drop=True)
    return df.rename(columns={'Product name': 'desc', 'Brand': 'brand'})

def read_grandmall(path):
    df = pd.read_excel(path, usecols=['Product name', 'Brand'])
    df = df.dropna(subset=['Product name']).drop_duplicates(subset=['Product name']).reset_index(drop=True)
    return df.rename(columns={'Product name': 'desc', 'Brand': 'brand'})

def read_qnie(path):
    df = pd.read_excel(path, header=3)
    desc_col = df.columns[4]
    df = df[['Brand', desc_col]].copy()
    df.columns = ['brand', 'desc']
    df = df.dropna(subset=['desc']).drop_duplicates(subset=['desc']).reset_index(drop=True)
    return df[['desc', 'brand']]

def read_rawspar(path):
    df = pd.read_excel(path, usecols=['Product name', 'Retail Article Brand Description Text'])
    df = df.rename(columns={'Product name': 'family', 'Retail Article Brand Description Text': 'brand'})
    df = df.dropna(subset=['brand']).drop_duplicates(subset=['family', 'brand']).reset_index(drop=True)
    df['desc'] = df['brand'].astype(str).str.upper() + ' ' + df['family'].astype(str).str.upper()
    return df[['desc', 'brand']]

def read_talabat(path):
    df = pd.read_excel(path, usecols=['Product name', 'Brand'])
    df = df.dropna(subset=['Product name']).drop_duplicates(subset=['Product name']).reset_index(drop=True)
    return df.rename(columns={'Product name': 'desc', 'Brand': 'brand'})

READERS = {
    'read_almeera':   read_almeera,
    'read_c4':        read_c4,
    'read_grandmall': read_grandmall,
    'read_qnie':      read_qnie,
    'read_rawspar':   read_rawspar,
    'read_talabat':   read_talabat,
}


# ===========================================================================
# MATCH ONE CUSTOMER FILE
# ===========================================================================

def match_customer(
    cust_cfg, src_name, src_weight, src_kind, src_pack, src_brands,
    src_descraw, src_mgrp, src_buckets, all_src_idxs,
    known_src_brands, semantic,
):
    label     = cust_cfg['label']
    reader_fn = READERS[cust_cfg['reader']]
    t0        = time.time()

    print(f"\n{'='*60}")
    print(f"  Customer: {label}  ({cust_cfg['path']})")
    print(f"{'='*60}")

    cust = reader_fn(cust_cfg['path'])
    cust['source_file'] = label
    print(f"  Rows: {len(cust):,}")

    # Clean descriptions before parsing (NEW)
    cust['_cleaned'] = cust['desc'].apply(
        lambda x: clean_description(x) if isinstance(x, str) else ''
    )

    parsed_cust     = [parse_description(t) for t in cust['_cleaned']]
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

    # v3: NO brand pre-gate — every row goes through matching
    print(f"  v3: Brand pre-gate disabled — all rows go through matching")

    cust_emb = None
    if semantic is not None:
        cust_emb = semantic.encode_batch(cust_name)

    # Fuzzy pass
    scorer = _ensemble_score if ENSEMBLE_FUZZY else fuzz.token_set_ratio
    fuzzy_results = [None] * len(cust)

    cust_groups = defaultdict(list)
    for ci, (kind, band) in enumerate(zip(cust['_kind'], cust['_band'])):
        kind = None if (isinstance(kind, float) and pd.isna(kind)) else kind
        cust_groups[(kind, band)].append(ci)

    bcount = 0
    for (kind, band), cidxs in cust_groups.items():
        bcount += 1
        if band is None:
            cand_idxs = all_src_idxs[:]
        else:
            cand_idxs = []
            for d in (-1, 0, 1):
                cand_idxs.extend(src_buckets.get((kind, band + d), []))
            cand_idxs.extend(src_buckets.get((kind, None), []))
            cand_idxs.extend(src_buckets.get((None, None), []))

        if not cand_idxs:
            for ci in cidxs:
                fuzzy_results[ci] = (None, 0.0, 'no_candidates', 'Fuzzy')
            continue

        cand_names  = [src_name[i]   for i in cand_idxs]
        query_names = [cust_name[ci] for ci in cidxs]
        if not query_names or not cand_names:
            for ci in cidxs:
                fuzzy_results[ci] = (None, 0.0, 'empty_names', 'Fuzzy')
            continue

        mat = process.cdist(query_names, cand_names, scorer=scorer,
                            workers=-1, dtype=np.float32)
        for q_i, ci in enumerate(cidxs):
            best_si, best_sc, reason = _apply_filters(
                mat[q_i].astype(np.float64), cand_idxs,
                cust_weight[ci], cust_kind[ci], cust_pack[ci], cust_brands[ci],
                src_weight, src_kind, src_pack, src_brands,
            )
            fuzzy_results[ci] = (best_si, best_sc, reason, 'Fuzzy')

        if bcount % 20 == 0 or bcount == len(cust_groups):
            print(f"  fuzzy bucket {bcount}/{len(cust_groups)}  ({time.time()-t0:.1f}s)")

    def _needs_rescue(r):
        if r is None: return True
        _, sc, reason, _ = r
        return reason in ('no_candidates', 'empty_names') \
               or sc < MIN_CONFIDENCE

    needs_semantic = [ci for ci, r in enumerate(fuzzy_results) if _needs_rescue(r)]
    print(f"  Fuzzy matched: {sum(1 for r in fuzzy_results if r and r[1]>=MIN_CONFIDENCE):,}"
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
            )
            final_results[ci] = (best_si, best_sc, reason, 'Semantic')
            if count % 100 == 0 or count == len(needs_semantic):
                print(f"  semantic {count}/{len(needs_semantic)}  ({time.time()-t0:.1f}s)")

    # Build rows
    matched_rows        = []
    unmatched_rows      = []
    alias_dict          = {}
    matched_src_indices = set()

    for ci in range(len(cust)):
        cust_desc_raw  = cust.at[ci, 'desc']
        cust_brand_raw = cust.at[ci, 'brand']
        cust_wt        = cust_weight[ci]
        cust_pk        = cust_pack[ci]
        result         = final_results[ci]

        base = {
            'Source File':            label,
            'Customer Description':   cust_desc_raw,
            'Customer Brand':         cust_brand_raw,
            'Customer Weight g/ml':   cust_wt,
            'Customer Pack':          cust_pk,
        }

        if result is None:
            unmatched_rows.append({**base, 'Reason': 'no_result', 'Best Score': 0})
            continue

        src_i, conf, reason, method = result

        if src_i is None or reason in ('below_threshold', 'no_candidates', 'empty_names'):
            unmatched_rows.append({**base, 'Reason': reason, 'Best Score': int(round(conf))})
            continue

        conf_int = int(round(conf))
        if conf_int < MIN_CONFIDENCE:
            unmatched_rows.append({**base, 'Reason': 'below_threshold', 'Best Score': conf_int})
            continue

        src_wt        = src_weight[src_i]
        src_pk        = src_pack[src_i]
        src_brand_raw = src_mgrp[src_i]
        blabel        = brand_label(src_brand_raw, cust_brand_raw, BRAND_FUZZY_THRESH, alias_dict)
        status        = 'Matched (High)' if conf_int >= 85 else 'Matched (Medium)'

        # Tag whether brand was a factor
        brand_note = ''
        if reason == 'brand_mismatch_but_matched':
            brand_note = 'Desc-only match'
        elif blabel in ('Exact', 'Fuzzy'):
            brand_note = 'Brand confirmed'
        else:
            brand_note = blabel

        matched_src_indices.add(src_i)

        matched_rows.append({
            'Source File':            label,
            'Our Brand (mgrp_descr)': src_brand_raw,
            'Our Name (parsed)':      src_name[src_i],
            'Our Weight g/ml':        src_wt,
            'Our Pack':               src_pk,
            'Our Description':        src_descraw[src_i],
            'Customer Description':   cust_desc_raw,
            'Customer Brand':         cust_brand_raw,
            'Customer Name (parsed)': cust_name[ci],
            'Customer Weight g/ml':   cust_wt,
            'Customer Pack':          cust_pk,
            'Match Status':           status,
            'Confidence Score':       conf_int,
            'Method':                 method,
            'Brand Match':            blabel,
            'Brand Note':             brand_note,      # NEW column
            'Weight Match':           weight_match_label(src_wt, cust_wt),
        })

    if SAMPLE_SIZE and len(matched_rows) > SAMPLE_SIZE:
        matched_rows = matched_rows[:SAMPLE_SIZE]

    desc_only = sum(1 for r in matched_rows if r.get('Brand Note') == 'Desc-only match')
    print(f"  → Matched: {len(matched_rows):,}  (of which desc-only brand: {desc_only:,})"
          f"  Unmatched: {len(unmatched_rows):,}  ({time.time()-t0:.1f}s)")

    return matched_rows, unmatched_rows, alias_dict, matched_src_indices


# ===========================================================================
# WRITE ALL RESULTS
# ===========================================================================

def write_all_results(
    all_matched, all_unmatched, alias_dict, per_file,
    total_src_rows, all_matched_src_indices, src_descraw,
):
    matched_df   = pd.DataFrame(all_matched)
    unmatched_df = pd.DataFrame(all_unmatched)

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
            pd.DataFrame(m_rows).to_excel(
                writer, sheet_name=f"{label} Matched"[:31], index=False)
            if u_rows:
                pd.DataFrame(u_rows).to_excel(
                    writer, sheet_name=f"{label} Unmatched"[:31], index=False)

        # Summary sheet
        rows = []
        rows.append({'Metric': '─── PER CUSTOMER FILE ───', 'Value': '', 'Detail': ''})
        rows.append({'Metric': '', 'Value': '', 'Detail': ''})

        for label, (m_rows, u_rows, src_idx) in per_file.items():
            t         = len(m_rows) + len(u_rows)
            match_pct = f"{len(m_rows)/t*100:.1f}%" if t else '0%'
            unmatch_pct = f"{len(u_rows)/t*100:.1f}%" if t else '0%'
            high      = sum(1 for r in m_rows if r['Match Status'] == 'Matched (High)')
            med       = sum(1 for r in m_rows if r['Match Status'] == 'Matched (Medium)')
            fuzzy     = sum(1 for r in m_rows if r['Method'] == 'Fuzzy')
            sem       = sum(1 for r in m_rows if r['Method'] == 'Semantic')
            desc_only = sum(1 for r in m_rows if r.get('Brand Note') == 'Desc-only match')
            cov_n     = len(src_idx)
            cov_p     = f"{cov_n / total_src_rows * 100:.1f}%" if total_src_rows else '0%'

            rows += [
                {'Metric': f'[{label}] Customer rows',              'Value': t,           'Detail': ''},
                {'Metric': f'[{label}] Matched',                    'Value': len(m_rows), 'Detail': match_pct},
                {'Metric': f'[{label}] Unmatched',                  'Value': len(u_rows), 'Detail': unmatch_pct},
                {'Metric': f'[{label}] High confidence (≥85)',      'Value': high,        'Detail': f"{high/len(m_rows)*100:.1f}% of matched" if m_rows else ''},
                {'Metric': f'[{label}] Med confidence (75-84)',     'Value': med,         'Detail': f"{med/len(m_rows)*100:.1f}% of matched" if m_rows else ''},
                {'Metric': f'[{label}] via Fuzzy',                  'Value': fuzzy,       'Detail': ''},
                {'Metric': f'[{label}] via Semantic',               'Value': sem,         'Detail': ''},
                {'Metric': f'[{label}] Desc-only brand matches',    'Value': desc_only,   'Detail': 'Brand mismatched but description matched'},
                {'Metric': f'[{label}] Source SKUs covered',        'Value': cov_n,       'Detail': f"{cov_p} of our {total_src_rows:,} source SKUs"},
                {'Metric': '', 'Value': '', 'Detail': ''},
            ]

        rows.append({'Metric': '─── COMBINED TOTALS (ALL FILES) ───', 'Value': '', 'Detail': ''})
        rows.append({'Metric': '', 'Value': '', 'Detail': ''})
        total_desc_only = sum(1 for r in all_matched if r.get('Brand Note') == 'Desc-only match')
        rows += [
            {'Metric': 'Total customer rows',           'Value': total_cust_rows,    'Detail': ''},
            {'Metric': 'Total matched',                 'Value': len(all_matched),   'Detail': f"{len(all_matched)/total_cust_rows*100:.1f}% of all customer rows" if total_cust_rows else ''},
            {'Metric': 'Total desc-only brand matches', 'Value': total_desc_only,    'Detail': 'Brand mismatched but description still matched'},
            {'Metric': 'Total unmatched',               'Value': len(all_unmatched), 'Detail': f"{len(all_unmatched)/total_cust_rows*100:.1f}% of all customer rows" if total_cust_rows else ''},
            {'Metric': '', 'Value': '', 'Detail': ''},
        ]

        rows.append({'Metric': '─── SOURCE CATALOGUE COVERAGE ───', 'Value': '', 'Detail': ''})
        rows.append({'Metric': '', 'Value': '', 'Detail': ''})
        rows += [
            {'Metric': 'Total source SKUs (our file)',   'Value': total_src_rows,    'Detail': '100%'},
            {'Metric': 'Source SKUs matched ≥1 time',   'Value': src_covered_count, 'Detail': f"{src_covered_pct:.1f}%  ← covered across all customer files"},
            {'Metric': 'Source SKUs never matched',     'Value': src_not_covered,   'Detail': f"{100 - src_covered_pct:.1f}%  ← not seen in any customer file"},
        ]

        pd.DataFrame(rows).to_excel(writer, sheet_name='Summary', index=False)

        # Source Never Matched sheet
        never_idx = sorted(set(range(total_src_rows)) - all_matched_src_indices)
        if never_idx:
            pd.DataFrame({
                'Source Index':    never_idx,
                'Our Description': [src_descraw[i] for i in never_idx],
            }).to_excel(writer, sheet_name='Source Never Matched', index=False)

        # Brand alias dictionary
        pd.DataFrame([
            {'Source Brand': s, 'Customer Brand': v['Customer Brand'], 'Fuzzy Score': v['Fuzzy Score']}
            for s, v in sorted(alias_dict.items())
        ]).to_excel(writer, sheet_name='Brand Alias Dict', index=False)

    print(f"  Total matched  : {len(matched_df):,}")
    print(f"  Total unmatched: {len(unmatched_df):,}")
    print(f"  Source coverage: {src_covered_count:,} / {total_src_rows:,} ({src_covered_pct:.1f}%)")


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    t_total = time.time()

    if not Path(OUR_FILE).exists():
        print(f"ERROR: source file not found: {OUR_FILE}"); sys.exit(1)
    missing = [c for c in CUSTOMER_FILES if not Path(c['path']).exists()]
    if missing:
        for m in missing:
            print(f"ERROR: customer file not found: {m['path']} ({m['label']})")
        sys.exit(1)
    bad = [c for c in CUSTOMER_FILES if c['reader'] not in READERS]
    if bad:
        for u in bad:
            print(f"ERROR: unknown reader '{u['reader']}' for {u['label']}.")
        sys.exit(1)

    print("Loading source catalogue …")
    src = pd.read_csv(OUR_FILE, low_memory=False)
    src = src.dropna(subset=['material_desc']).reset_index(drop=True)

    # Clean source descriptions too (NEW)
    src['_cleaned'] = src['material_desc'].apply(
        lambda x: clean_description(x) if isinstance(x, str) else ''
    )

    parsed_src     = [parse_description(t) for t in src['_cleaned']]
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
    all_src_idxs = list(range(total_src_rows))

    semantic = None
    if USE_SEMANTIC:
        print("\nBuilding semantic index …")
        try:
            semantic = SemanticMatcher(src_name)
        except ImportError as e:
            print(f"  WARNING: {e}\n  Continuing with fuzzy-only.")

    all_matched             = []
    all_unmatched           = []
    all_aliases             = {}
    per_file                = {}
    all_matched_src_indices = set()

    for cust_cfg in CUSTOMER_FILES:
        m_rows, u_rows, aliases, src_idx_set = match_customer(
            cust_cfg,
            src_name, src_weight, src_kind, src_pack, src_brands,
            src_descraw, src_mgrp, src_buckets, all_src_idxs,
            known_src_brands, semantic,
        )
        all_matched.extend(m_rows)
        all_unmatched.extend(u_rows)
        all_aliases.update(aliases)
        per_file[cust_cfg['label']] = (m_rows, u_rows, src_idx_set)
        all_matched_src_indices |= src_idx_set

    write_all_results(
        all_matched, all_unmatched, all_aliases, per_file,
        total_src_rows, all_matched_src_indices, src_descraw,
    )

    print(f"\nTotal time: {time.time()-t_total:.1f}s")
    print(f"Saved → {OUTPUT_FILE}")


if __name__ == '__main__':
    main()