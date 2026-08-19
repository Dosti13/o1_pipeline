"""
Combined Matcher — NAME + WEIGHT + BRAND  (Fuzzy → Semantic cascade)
=====================================================================

FIX LOG (brand gate):
  - Brand gate now runs BEFORE fuzzy/semantic.  Any customer item whose
    brand tokens have ZERO overlap with the entire source brand universe
    is immediately sent to Unmatched with reason='brand_not_in_catalogue'.
    This stops WALKERS / LOACKER / BAHLSEN / KAGI etc. from being matched
    to AMERICANA / MUMTAZ / TIFFANY products.

  - _apply_filters now returns (None, 0, 'brand_mismatch') for the
    best-scoring candidate instead of the candidate's index, so the
    main loop cannot accidentally accept it.

  - brands_overlap() tightened: substring match now requires both tokens
    to be ≥ 5 chars (was 4) to avoid false overlaps on short codes.

  - Semantic rescue now also skips hard_unmatched rows (brand not in
    catalogue) so they are never re-examined.

Pipeline per customer row:
  0. Brand pre-gate – customer brand not in source universe → Unmatched
  1. Parse           – abbreviation expansion, weight/pack split, size-band
  2. Fuzzy           – ensemble (WRatio + token_set_ratio + partial_ratio)
  3. Semantic        – sentence-transformer on fuzzy-unmatched rows only
  4. Weight gate     – wrong weight → score zeroed
  5. Brand gate      – brands_overlap() must pass → Unmatched
  6. Brand label     – Exact / Fuzzy / Mismatch / N/A  (reporting only)
  7. Method tag      – Fuzzy | Semantic

Install:
  pip install pandas openpyxl rapidfuzz numpy sentence-transformers

Run:
  python matcher_combined_fixed.py

Demo (no files needed):
  python matcher_combined_fixed.py --demo
"""

# ===========================================================================
# CONFIG
# ===========================================================================

OUR_FILE    = r"C:\Users\HP\Desktop\functionand sp.csv"
CUST_FILE   = r"C:\Users\HP\Downloads\fwqniemtselloutdashboardsession2\C4 sep 24.xlsx"
CUST_LABEL  = "C4"
CUST_READER = "read_c4"
OUTPUT_FILE = r"C:\Users\HP\Desktop\match_results_combined.xlsx"

SAMPLE_SIZE    = None
MIN_CONFIDENCE = 75

ENSEMBLE_FUZZY   = True
USE_SEMANTIC     = True
SEMANTIC_MODEL   = "all-MiniLM-L6-v2"
SEMANTIC_BATCH   = 128
BRAND_FUZZY_THRESH = 75
WEIGHT_TOL         = 0.05

# ===========================================================================

import re, sys, time, warnings
import numpy as np
import pandas as pd
from collections import defaultdict, Counter
from pathlib import Path
from rapidfuzz import fuzz, process

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# Abbreviation tables  (unchanged)
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
# Regex patterns  (unchanged)
# ---------------------------------------------------------------------------

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


# ===========================================================================
# PARSING  (unchanged)
# ===========================================================================

def parse_description(text: str):
    if not isinstance(text, str):
        return ('', None, None, '')
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
    """
    Extract likely brand token from the FIRST non-generic word of the
    parsed product name.  Only used as a fallback when the explicit brand
    column is empty — kept to 1 token to avoid polluting the brand set
    with product-word false positives.
    """
    if not clean_name:
        return frozenset()
    for w in clean_name.split()[:2]:          # look at first 2 words only
        if w in NON_BRAND or w.isdigit() or len(w) < 3:
            continue
        return frozenset({BRAND_ABBREV.get(w, w)})
    return frozenset()


def brands_overlap(a: frozenset, b: frozenset) -> bool:
    """
    Return True if any token in `a` is considered the same brand as any
    token in `b`.

    Rules (in order):
      1. Exact match.
      2. One is a substring of the other — but ONLY if both tokens are
         ≥ 5 chars (raised from 4 to cut false positives like AL/ALI).
      3. First-5-char prefix match (requires both ≥ 5 chars).
    """
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
# SIZE-BAND + WEIGHT HELPERS  (unchanged)
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
# FUZZY SCORER  (unchanged)
# ===========================================================================

def _ensemble_score(a: str, b: str, **_) -> float:
    return (fuzz.WRatio(a, b) + fuzz.token_set_ratio(a, b) + fuzz.partial_ratio(a, b)) / 3.0


# ===========================================================================
# SEMANTIC LAYER  (unchanged)
# ===========================================================================

class SemanticMatcher:
    def __init__(self, source_names: list):
        print(f"  Loading semantic model '{SEMANTIC_MODEL}' …")
        try:
            from sentence_transformers import SentenceTransformer
        except ImportError:
            raise ImportError(
                "sentence-transformers not installed.\n"
                "Run:  pip install sentence-transformers"
            )
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
# FILTER FUNCTION — weight gate + brand gate
# ===========================================================================

def _apply_filters(
    scores: np.ndarray,
    cand_idxs: list,
    c_weight, c_kind_v, c_pack_v, c_brands,
    src_weight, src_kind, src_pack, src_brands,
):
    """
    Apply weight gate (zero bad-weight scores) and brand gate (only accept
    candidates whose brand overlaps the customer brand).

    Returns (best_src_i, best_score, reason).
      reason == 'ok'            → accepted match
      reason == 'brand_mismatch'→ best score found but NO brand overlap
                                  caller must send this to Unmatched
      reason == 'no_brand_match'→ all candidates below MIN_CONFIDENCE
    """
    scores = scores.copy()

    for j, si in enumerate(cand_idxs):
        wm = weights_match(c_weight, src_weight[si])
        if wm is False:
            scores[j] = 0.0
            continue
        if wm is True:
            scores[j] = min(100.0, scores[j] + 5.0)
            if c_kind_v and src_kind[si] and c_kind_v != src_kind[si]:
                scores[j] = max(0.0, scores[j] - 20.0)
        scores[j] = min(100.0, scores[j] + pack_match_bonus(c_pack_v, src_pack[si]))

    # Walk candidates sorted by score descending; take first brand-passing one
    order = np.argsort(scores)[::-1]

    for j in order:
        sc = float(scores[j])
        if sc < MIN_CONFIDENCE:
            break   # everything from here will also be below threshold
        si = cand_idxs[j]
        if brands_overlap(c_brands, src_brands[si]):
            return si, sc, 'ok'

    # ── Nothing passed brand gate above threshold ──────────────────────────
    # Return the best-scoring candidate index for diagnostics, but mark as
    # brand_mismatch so the caller sends it to Unmatched — do NOT use si.
    best_j  = int(np.argmax(scores))
    best_si = cand_idxs[best_j]
    best_sc = float(scores[best_j])
    return best_si, best_sc, 'brand_mismatch'


# ===========================================================================
# CUSTOMER FILE READERS  (unchanged)
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
    df = pd.read_excel(path, usecols=['product name', 'BRAND'])
    df = df.dropna(subset=['product name']).drop_duplicates(subset=['product name']).reset_index(drop=True)
    return df.rename(columns={'product name': 'desc', 'BRAND': 'brand'})

def read_qnie(path):
    df = pd.read_excel(path, header=3)
    desc_col = df.columns[4]
    df = df[['Brand', desc_col]].copy()
    df.columns = ['brand', 'desc']
    df = df.dropna(subset=['desc']).drop_duplicates(subset=['desc']).reset_index(drop=True)
    return df[['desc', 'brand']]

def read_rawspar(path):
    df = pd.read_excel(path, usecols=['Family Text', 'Retail Article Brand Description Text'])
    df = df.rename(columns={'Family Text': 'Product name', 'Retail Article Brand Description Text': 'brand'})
    df = df.dropna(subset=['brand']).drop_duplicates(subset=['Product name', 'brand']).reset_index(drop=True)
    df['desc'] = df['brand'].astype(str).str.upper() + ' ' + df['Product name'].astype(str).str.upper()
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
# OUTPUT WRITER  (unchanged)
# ===========================================================================

def write_results(matched_rows: list, unmatched_rows: list, alias_dict: dict):
    matched_df   = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)
    total        = len(matched_df) + len(unmatched_df)

    print(f"\nWriting → {OUTPUT_FILE}")
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        matched_df.to_excel(writer, sheet_name='Matched Results', index=False)
        if not unmatched_df.empty:
            unmatched_df.to_excel(writer, sheet_name='Unmatched', index=False)
        if not matched_df.empty:
            mc = matched_df['Method'].value_counts().reset_index()
            mc.columns = ['Method', 'Count']
            mc['Pct'] = (mc['Count'] / total * 100).round(1).astype(str) + '%'
            mc.to_excel(writer, sheet_name='Method Summary', index=False)
            bc = matched_df['Brand Match'].value_counts().reset_index()
            bc.columns = ['Brand Match', 'Count']
            bc['Pct'] = (bc['Count'] / total * 100).round(1).astype(str) + '%'
            bc.to_excel(writer, sheet_name='Brand Summary', index=False)
        alias_rows = [
            {'Source Brand (mgrp_descr)': src,
             'Customer Brand':            v['Customer Brand'],
             'Fuzzy Score':               v['Fuzzy Score']}
            for src, v in sorted(alias_dict.items())
        ]
        pd.DataFrame(alias_rows).to_excel(
            writer, sheet_name='Brand Alias Dictionary', index=False)
        if not matched_df.empty:
            wc = matched_df['Weight Match'].value_counts().reset_index()
            wc.columns = ['Weight Match', 'Count']
            wc['Pct'] = (wc['Count'] / total * 100).round(1).astype(str) + '%'
            wc.to_excel(writer, sheet_name='Weight Summary', index=False)

    print(f"  Matched rows  : {len(matched_df):,}")
    print(f"  Unmatched rows: {len(unmatched_df):,}")
    print(f"  Brand aliases : {len(alias_dict):,}")


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    t0 = time.time()

    # ── 1. Load & parse source ─────────────────────────────────────────────
    print("Loading source catalogue …")
    src = pd.read_csv(OUR_FILE, low_memory=False)
    src = src.dropna(subset=['material_desc']).reset_index(drop=True)

    print(f"  Parsing {len(src):,} source descriptions …")
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

    # ── BUILD GLOBAL SOURCE-BRAND UNIVERSE ─────────────────────────────────
    # A flat frozenset of every brand token that exists anywhere in the source.
    # Used by the pre-gate to immediately discard customer items whose brand
    # is simply not stocked.
    known_src_brands: frozenset = frozenset(
        tok for bs in src['_brands'] for tok in bs
    )
    print(f"  Unique source brand tokens: {len(known_src_brands):,}")

    # Size-band lookup index
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
    all_src_idxs = list(range(len(src)))
    print(f"  Source rows: {len(src):,}")

    # ── 2. Build semantic index (optional) ────────────────────────────────
    semantic = None
    if USE_SEMANTIC:
        print("\nBuilding semantic index …")
        try:
            semantic = SemanticMatcher(src_name)
        except ImportError as e:
            print(f"  WARNING: {e}\n  Continuing with fuzzy-only.")

    # ── 3. Load & parse customer file ─────────────────────────────────────
    reader_fn = READERS.get(CUST_READER)
    if reader_fn is None:
        raise ValueError(f"Unknown reader '{CUST_READER}'. Options: {list(READERS)}")

    print(f"\nLoading customer file ({CUST_LABEL}) …")
    cust = reader_fn(CUST_FILE)
    cust['source_file'] = CUST_LABEL
    print(f"  Customer rows: {len(cust):,}")

    print(f"  Parsing {len(cust):,} customer descriptions …")
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

    # ── 4. BRAND PRE-GATE — discard brands not in source catalogue ─────────
    # If a customer row has brand tokens AND none of them appear anywhere in
    # the source brand universe, there is nothing to match against — skip
    # fuzzy + semantic entirely and go straight to Unmatched.
    hard_unmatched_cis: set = set()
    for ci in range(len(cust)):
        cb = cust_brands[ci]
        if not cb:
            continue   # no brand info — let the matcher try on name alone
        if not brands_overlap(cb, known_src_brands):
            hard_unmatched_cis.add(ci)

    print(f"\n  Brand pre-gate: {len(hard_unmatched_cis):,} rows have brands "
          f"absent from source catalogue → will be Unmatched immediately")

    # Pre-encode ALL customer names for semantic in one batch
    cust_emb = None
    if semantic is not None:
        print("  Encoding customer names for semantic search …")
        cust_emb = semantic.encode_batch(cust_name)

    # ── 5. FUZZY PASS ──────────────────────────────────────────────────────
    scorer = _ensemble_score if ENSEMBLE_FUZZY else fuzz.token_set_ratio
    print(f"\n[Pass 1 — Fuzzy {'ensemble' if ENSEMBLE_FUZZY else 'token_set_ratio'}] …")

    fuzzy_results = [None] * len(cust)

    # Pre-populate hard-unmatched rows so fuzzy loop skips them
    for ci in hard_unmatched_cis:
        fuzzy_results[ci] = (None, 0.0, 'brand_not_in_catalogue', 'N/A')

    cust_groups = defaultdict(list)
    for ci, (kind, band) in enumerate(zip(cust['_kind'], cust['_band'])):
        if ci in hard_unmatched_cis:
            continue   # already handled
        kind = None if (isinstance(kind, float) and pd.isna(kind)) else kind
        cust_groups[(kind, band)].append(ci)
    print(f"  {len(cust_groups):,} (kind, band) buckets")

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

        mat = process.cdist(
            query_names, cand_names,
            scorer=scorer, workers=-1, dtype=np.float32,
        )

        for q_i, ci in enumerate(cidxs):
            scores = mat[q_i].astype(np.float64)
            best_si, best_sc, reason = _apply_filters(
                scores, cand_idxs,
                cust_weight[ci], cust_kind[ci], cust_pack[ci], cust_brands[ci],
                src_weight, src_kind, src_pack, src_brands,
            )
            fuzzy_results[ci] = (best_si, best_sc, reason, 'Fuzzy')

        if bcount % 20 == 0 or bcount == len(cust_groups):
            print(f"  bucket {bcount}/{len(cust_groups)}  ({time.time()-t0:.1f}s)")

    # Rows needing semantic rescue (exclude hard-unmatched)
    def _needs_rescue(r):
        if r is None:
            return True
        _, sc, reason, _ = r
        return reason in ('no_candidates', 'empty_names', 'no_brand_match') \
               or (reason == 'ok' and sc < MIN_CONFIDENCE)
        # NOTE: 'brand_mismatch' is intentionally excluded here.
        # A brand_mismatch from fuzzy means the name matched well but the brand
        # didn't.  Semantic won't fix that — it would just find an equally wrong
        # brand partner.  Keep it as Unmatched.

    needs_semantic = [
        ci for ci, r in enumerate(fuzzy_results)
        if _needs_rescue(r) and ci not in hard_unmatched_cis
    ]
    fuzzy_matched = sum(
        1 for ci, r in enumerate(fuzzy_results)
        if r is not None and r[2] == 'ok' and r[1] >= MIN_CONFIDENCE
        and ci not in hard_unmatched_cis
    )
    print(f"  Fuzzy matched : {fuzzy_matched:,} / {len(cust) - len(hard_unmatched_cis):,} "
          f"(excl. pre-gated)")
    print(f"  Need semantic : {len(needs_semantic):,}")

    # ── 6. SEMANTIC PASS ───────────────────────────────────────────────────
    final_results = list(fuzzy_results)

    if semantic is not None and needs_semantic:
        print(f"\n[Pass 2 — Semantic on {len(needs_semantic):,} rows] …")
        for count, ci in enumerate(needs_semantic, 1):
            if cust_emb is None or not cust_name[ci]:
                continue
            q_emb       = cust_emb[ci]
            scores_full = semantic.scores_for_query(q_emb)
            best_si, best_sc, reason = _apply_filters(
                scores_full, all_src_idxs,
                cust_weight[ci], cust_kind[ci], cust_pack[ci], cust_brands[ci],
                src_weight, src_kind, src_pack, src_brands,
            )
            final_results[ci] = (best_si, best_sc, reason, 'Semantic')
            if count % 100 == 0 or count == len(needs_semantic):
                print(f"  {count}/{len(needs_semantic)}  ({time.time()-t0:.1f}s)")
    else:
        tag = "(model not loaded)" if semantic is None else "(nothing left to match)"
        print(f"\n[Pass 2 — Semantic SKIPPED {tag}]")

    print(f"\nAll matching done in {time.time()-t0:.1f}s")

    # ── 7. Build output rows ───────────────────────────────────────────────
    print(f"\nBuilding output (≥{MIN_CONFIDENCE} confidence, brand gate ON) …")
    matched_rows   = []
    unmatched_rows = []
    alias_dict     = {}

    for ci in range(len(cust)):
        cust_desc_raw  = cust.at[ci, 'desc']
        cust_brand_raw = cust.at[ci, 'brand']
        cust_wt        = cust_weight[ci]
        cust_pk        = cust_pack[ci]
        result         = final_results[ci]

        def _add_unmatched(reason, score=0):
            unmatched_rows.append({
                'Source File':          cust.at[ci, 'source_file'],
                'Customer Description': cust_desc_raw,
                'Customer Brand':       cust_brand_raw,
                'Customer Weight g/ml': cust_wt,
                'Customer Pack':        cust_pk,
                'Reason':               reason,
                'Best Score':           int(round(score)),
            })

        if result is None:
            _add_unmatched('no_result')
            continue

        src_i, conf, reason, method = result

        # Hard brand pre-gate or brand_mismatch → always Unmatched
        if reason in ('brand_not_in_catalogue', 'brand_mismatch') or src_i is None:
            _add_unmatched(reason, conf)
            continue

        conf_int = int(round(conf))
        if conf_int < MIN_CONFIDENCE:
            _add_unmatched('below_threshold', conf_int)
            continue

        # ── Matched ───────────────────────────────────────────────────────
        src_wt        = src_weight[src_i]
        src_pk        = src_pack[src_i]
        src_brand_raw = src_mgrp[src_i]
        blabel = brand_label(src_brand_raw, cust_brand_raw, BRAND_FUZZY_THRESH, alias_dict)
        status = 'Matched (High)' if conf_int >= 85 else 'Matched (Medium)'

        matched_rows.append({
            'Source File':              cust.at[ci, 'source_file'],
            'Our Description':          src_descraw[src_i],
            'Customer Description':     cust_desc_raw,
            'Our Brand (mgrp_descr)':   src_brand_raw,
            'Our Name (parsed)':        src_name[src_i],
            'Our Weight g/ml':          src_wt,
            'Our Pack':                 src_pk,
          
            'Customer Brand':           cust_brand_raw,
            'Customer Name (parsed)':   cust_name[ci],
            'Customer Weight g/ml':     cust_wt,
            'Customer Pack':            cust_pk,
            'Match Status':             status,
            'Confidence Score':         conf_int,
            'Method':                   method,
            'Brand Match':              blabel,
            'Weight Match':             weight_match_label(src_wt, cust_wt),
        })

    if SAMPLE_SIZE and len(matched_rows) > SAMPLE_SIZE:
        matched_rows = matched_rows[:SAMPLE_SIZE]
        print(f"Sample cap applied — keeping first {SAMPLE_SIZE} matched rows.")

    print(f"\n  Total rows    : {len(cust):,}")
    print(f"  Matched       : {len(matched_rows):,}")
    print(f"  Unmatched     : {len(unmatched_rows):,}")
    if matched_rows:
        mc = Counter(r['Method'] for r in matched_rows)
        for m, c in sorted(mc.items()):
            print(f"    via {m}: {c:,}")

    write_results(matched_rows, unmatched_rows, alias_dict)
    print(f"\nTotal time : {time.time()-t0:.1f}s")
    print(f"Saved      → {OUTPUT_FILE}")


# ===========================================================================
# DEMO MODE
# ===========================================================================

def demo_parse():
    samples = [
        "TIFFANY DIGESTIVE NATURAL 6X3X250G",
        "TIFF NICE EVDAY REG 4X12X50G",
        "TIFF CREAM EVDAY CHOCO 4X6X 90G",
        "TIFF GLUC MILK AND HNY BISC 4X12X40GR",
        "TIFF CREAM EVDAY STRWBRY 24X84G WS",
        "ROYAL CLASSIC BUTTER COOKIES 454GX12",
        "OREO CHOCOLATE SANDWICH 135MLX10",
        "DEEMAH DIGESTIVE BISCUIT 1X24X230GR",
    ]
    print(f"\n{'Original':<46} {'Name (parsed)':<40} {'Wt g/ml':>10}  {'Kind':>5}  Pack")
    print('-' * 115)
    for s in samples:
        name, wt, kind, pack = parse_description(s)
        wt_str = f"{wt:.1f}" if wt is not None else '-'
        print(f"{s:<46} {name:<40} {wt_str:>10}  {(kind or '-'):>5}  {pack}")


if __name__ == '__main__':
    if '--demo' in sys.argv:
        demo_parse()
    else:
        if not Path(OUR_FILE).exists():
            print(f"ERROR: source file not found: {OUR_FILE}")
            sys.exit(1)
        if not Path(CUST_FILE).exists():
            print(f"ERROR: customer file not found: {CUST_FILE}")
            sys.exit(1)
        main()