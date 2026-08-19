"""
LULU CATEGORY MATCHING PIPELINE — v8.0 (CSV-only, no database)
================================================================
Matches non-LULU customer material names (AL_MEERA, C4, TALABAT, GRAND_MALL,
etc.) against LULU's benchmark category taxonomy.

WHAT CHANGED FROM v7.1
-----------------------
1. ALL Postgres/psycopg2 code removed — no DB connection, no DB loaders,
   no DB inserts, no classify_db(). Pipeline is pure CSV in -> CSV out.
2. "Master" is now the LULU category list (lulu_categories_updated.csv,
   one column: `category`) instead of a SKU master table.
3. "Customer" input is a single CSV/XLSX file with material-level rows
   (material_code, material_name, brand, source_customer,
   customer_category, ...) instead of multiple per-customer DB tables.
4. Barcode matching removed (categories have no barcodes).
5. Brand-strict-rejection logic removed — brand doesn't determine a
   product's category, so gating on brand mismatch was actively wrong
   here and would have rejected valid matches. Kept normalize_text()'s
   noise-stripping (units, pack sizes, stopwords) since that still helps
   both category text and material text line up semantically.
6. Matching signal is now: (a) exact normalized customer_category ==
   LULU category, else (b) AI embedding similarity of material_name
   against every LULU category (FAISS), blended with (c) rapidfuzz
   text similarity, mirroring the ai+rf blend style from v7.1.

USAGE
-----
Set CSV_LULU_CATEGORIES / CSV_CUSTOMER_INPUT / CSV_OUTPUT via env vars,
or just edit CONFIG below, then: python lulu_category_matching_pipeline.py
"""

import os
import re
import sys
import logging
import datetime as dt
from pathlib import Path

import numpy as np


def _ensure(pkg, import_as=None):
    name = import_as or pkg
    try:
        __import__(name)
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg,
                               "--quiet", "--break-system-packages"])


_ensure("pandas")
_ensure("openpyxl")
_ensure("sentence-transformers", "sentence_transformers")
_ensure("faiss-cpu", "faiss")
_ensure("rapidfuzz")

import pandas as pd
from sentence_transformers import SentenceTransformer
import faiss
from rapidfuzz import fuzz

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
CONFIG = {
    # Benchmark: LULU category taxonomy (one column, e.g. "category")

    # Customer rows to match against LULU categories (.csv or .xlsx)
    "csv_customer_input": r"C:\Users\HP\Desktop\categorymatch.csv",
    "csv_output_path": r"C:\Users\HP\Desktop\lulu_category_match_results.csv",
    

    "model_name": "all-MiniLM-L6-v2",
    "embedding_batch_size": 256,
    "faiss_top_k": 5,

    # Final blended score (0-1 scale) at/above which a row is MATCHED
    "match_threshold": 0.55,
    # Below this, flagged REVIEW instead of UNMATCHED (borderline zone)
    "review_threshold": 0.40,

    "ngram_range": (1, 2),
}

# --------------------------------------------------------------------------
# TEXT NORMALIZATION (kept from v7.1 — still useful for material_name and
# category text cleanup; brand-abbreviation/brand-gating logic removed)
# --------------------------------------------------------------------------
_STOPWORDS = frozenset({
    'the', 'a', 'an', 'and', 'or', 'of', 'in', 'for', 'with', 'by', 'from', 'to', 'at', 'on',
    'is', 'are', 'new', 'original', 'premium', 'special', 'classic', 'natural', 'fresh',
    'pure', 'extra', 'super', 'ultra', 'mini', 'big', 'large', 'small', 'best', 'real',
    'true', 'fine', 'frozen', 'chilled', 'hot', 'spicy', 'crispy', 'crunchy', 'regular',
    'light', 'full', 'fat', 'low', 'rich', 'deluxe', 'assorted', 'asstd', 'asst',
    'prm', 'prem', 'std', 'alu', 'tetra', 'cpd', 'chopped', 'whole', 'sliced', 'shredded',
    'diced', 'minced', 'bottle', 'jar', 'tub', 'pouch', 'family', 'traditional', 'artisan',
    'sugar', 'lean', 'skimmed', 'piece', 'stem', 'chunk', 'strip', 'sachet', 'squeeze',
    'spray', 'drinking', 'mineral', 'spring', 'purified', 'organic', 'bio', 'homestyle',
    'kids', 'junior', 'baby', 'adult', 'promo', 'offer', 'deal', 'halal', 'kosher',
    'vegan', 'vegetarian', 'gold', 'silver', 'platinum', 'bronze', 'smoked', 'roasted',
    'grilled', 'baked', 'boiled', 'cooked', 'marinated', 'seasoned', 'unseasoned', 'breaded',
})

_UNIT_PATTERNS = [
    (re.compile(r'\b(ltr|lt|litre|liter|liters|litres)\b'), 'l'),
    (re.compile(r'\b(grm|gms|gram|grams|gr|gm)\b'), 'g'),
    (re.compile(r'\b(mltr|milliliter|millilitre)\b'), 'ml'),
    (re.compile(r'\b(kilogram|kilograms)\b'), 'kg'),
    (re.compile(r'\b(ounce|ounces)\b'), 'oz'),
    (re.compile(r'\b(piece|pieces)\b'), 'pcs'),
    (re.compile(r'\b(packet|packets|pkt)\b'), 'pack'),
]

_PACK_RE = re.compile(r'\b(carton|ctn|case|pack|pcs?|units?|pieces?|promo|shrink|bundle|offer|sp|foc|eoe|eof|po)\b')
_MULTI_RE = re.compile(r'\b\d+[xX]\d+[xX]\d+[xX]?(\d+\.?\d*(?:ml|l|kg|g|oz|lb)?)\b')
_MULTI2_RE = re.compile(r'\b\d+[xX]\d+[xX](\d+\.?\d*(?:ml|l|kg|g|oz|lb)?)\b')
_SIZE_RE = re.compile(r'(\d+\.?\d*)\s*(ml|l|kg|g|oz|lb)\b')
_XQTY_RE = re.compile(r'\s*[xX]\s*\d+\b(?!\s*(?:ml|l|kg|g|oz|lb))')
_PCT_RE = re.compile(r'\d+%\s*(?:off|extra|free|discount)?\b')
_FREE_RE = re.compile(r'\(\d+\s*free\)|\b\d+\s*free\b')
_ALNUM_RE = re.compile(r'[^\w\s]')
_SPACE_RE = re.compile(r'\s+')


def normalize_text(text: str) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    t = text.strip().lower()
    for pat, repl in _UNIT_PATTERNS:
        t = pat.sub(repl, t)
    t = re.sub(r'(\d+)(m)\b(?!l)', r'\1ml', t)
    t = _MULTI_RE.sub(r'\1', t)
    t = _MULTI2_RE.sub(r'\1', t)
    t = _SIZE_RE.sub(r'\1\2', t)
    t = _XQTY_RE.sub(' ', t)
    t = _PACK_RE.sub('', t)
    t = re.sub(r'\b(ndw|foc|alu|tb|sl|cpd|tetra|ec|qnie)\b', '', t)
    t = _FREE_RE.sub('', t)
    t = _PCT_RE.sub('', t)
    t = re.sub(r'\b(\d+)s\b', r'\1', t)
    t = _ALNUM_RE.sub(' ', t)
    t = re.sub(r'([a-z])(\d)', r'\1 \2', t)
    t = re.sub(r'(\d)([a-z])', r'\1 \2', t)
    t = re.sub(r'\b(gr|gm)\b', 'g', t)
    tokens = [tok for tok in t.split() if tok not in _STOPWORDS]
    return _SPACE_RE.sub(' ', ' '.join(tokens)).strip()


def rapidfuzz_score(q_norm: str, cand_norm: str) -> float:
    if not q_norm or not cand_norm:
        return 0.0
    s_sort = fuzz.token_sort_ratio(q_norm, cand_norm) / 100.0
    s_set = fuzz.token_set_ratio(q_norm, cand_norm) / 100.0
    return 0.5 * s_sort + 0.5 * s_set


# --------------------------------------------------------------------------
# LOADERS (CSV / XLSX — no database anywhere)
# --------------------------------------------------------------------------
def load_lulu_categories(path: str) -> tuple:
    """Load the LULU benchmark category list. Returns (raws, norms)."""
    log.info(f"  Loading LULU category benchmark: {path}")
    df = pd.read_csv(path)
    cat_col = next((c for c in df.columns if 'categ' in c.lower()), df.columns[0])
    df = df.dropna(subset=[cat_col])
    raws = df[cat_col].astype(str).str.strip().tolist()
    raws = [r for r in raws if r]
    norms = [normalize_text(r) for r in raws]
    log.info(f"  LULU categories loaded: {len(raws):,} distinct benchmark categories")
    return raws, norms


def _read_any(path: str) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    print(ext)
    print(path)
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path)
    return pd.read_csv(path)


def load_customer_input(path: str) -> pd.DataFrame:
    """Load customer rows to be matched. Keeps all original columns and
    adds normalized text used for matching."""
    log.info(f"  Loading customer input: {path}")
    df = _read_any(path)
    # Drop fully-unnamed junk columns (e.g. stray "Unnamed: 10" from exports)
    df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]

    cols_lower = {c.lower(): c for c in df.columns}
    name_col = cols_lower.get("material_name") or next(
        (c for c in df.columns if "material" in c.lower() and "name" in c.lower()),
        None,
    )
    if name_col is None:
        # fallback: first free-text column
        name_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]

    cat_col = cols_lower.get("customer_category")

    df["_material_name"] = df[name_col].astype(str).fillna("")
    df["_customer_category"] = df[cat_col].astype(str).fillna("") if cat_col else ""
    df["_norm_material_name"] = df["_material_name"].apply(normalize_text)
    df["_norm_customer_category"] = df["_customer_category"].apply(normalize_text)

    log.info(f"  Customer input loaded: {len(df):,} rows "
             f"(name_col='{name_col}', category_col='{cat_col}')")
    return df


# --------------------------------------------------------------------------
# MATCHING ENGINE
# --------------------------------------------------------------------------
def build_ai_model():
    log.info(f"  Loading AI model: {CONFIG['model_name']}...")
    model = SentenceTransformer(CONFIG["model_name"])
    log.info(f"  Model loaded (dim={model.get_embedding_dimension()})")
    return model


def build_category_index(model, cat_norms: list):
    log.info(f"  Encoding {len(cat_norms):,} LULU category embeddings...")
    embeddings = model.encode(cat_norms, batch_size=CONFIG["embedding_batch_size"],
                              show_progress_bar=True, normalize_embeddings=True)
    dim = model.get_embedding_dimension()
    index = faiss.IndexFlatIP(dim)
    index.add(embeddings.astype(np.float32))
    return index


def match_customers_to_categories(df: pd.DataFrame, cat_raws: list, cat_norms: list,
                                   model, cat_index) -> pd.DataFrame:
    top_k = min(CONFIG["faiss_top_k"], len(cat_raws))
    emb_batch = CONFIG["embedding_batch_size"]
    match_thr = CONFIG["match_threshold"]
    review_thr = CONFIG["review_threshold"]
    cat_norm_to_idx = {n: i for i, n in enumerate(cat_norms)}

    n = len(df)
    matched_category = [None] * n
    scores = [0.0] * n
    methods = [None] * n
    statuses = [None] * n

    # ---- Pass 1: exact normalized customer_category == LULU category ----
    exact_mask = df["_norm_customer_category"].map(
        lambda x: x in cat_norm_to_idx and x != ""
    ).to_numpy()

    remaining_idx = []
    for i, is_exact in enumerate(exact_mask):
        if is_exact:
            idx = cat_norm_to_idx[df["_norm_customer_category"].iat[i]]
            matched_category[i] = cat_raws[idx]
            scores[i] = 100.0
            methods[i] = "exact_category"
            statuses[i] = "MATCHED"
        else:
            remaining_idx.append(i)

    log.info(f"  Pass 1 (exact category text): {n - len(remaining_idx):,} matched")

    # ---- Pass 2: AI embedding + rapidfuzz blend on material_name ----
    total_remaining = len(remaining_idx)
    counters = {"ai_rf": 0, "review": 0, "unmatched": 0}
    start_ts = dt.datetime.now()

    for batch_start in range(0, total_remaining, emb_batch):
        batch_idx = remaining_idx[batch_start: batch_start + emb_batch]
        batch_texts = [df["_norm_material_name"].iat[i] or df["_norm_customer_category"].iat[i]
                       for i in batch_idx]
        # rows with genuinely empty text can't be matched
        empty_flags = [not t for t in batch_texts]
        safe_texts = [t if t else "unknown" for t in batch_texts]

        q_emb = model.encode(safe_texts, batch_size=emb_batch,
                             normalize_embeddings=True, show_progress_bar=False).astype(np.float32)
        sc_mat, ix_mat = cat_index.search(q_emb, top_k)

        for j, row_i in enumerate(batch_idx):
            if empty_flags[j]:
                matched_category[row_i] = None
                scores[row_i] = 0.0
                methods[row_i] = "no_text"
                statuses[row_i] = "UNMATCHED"
                counters["unmatched"] += 1
                continue

            q_norm = batch_texts[j]
            cust_cat_norm = df["_norm_customer_category"].iat[row_i]

            best_final = 0.0
            best_cat_idx = int(ix_mat[j][0])

            for k in range(len(ix_mat[j])):
                c_idx = int(ix_mat[j][k])
                if c_idx < 0 or c_idx >= len(cat_norms):
                    continue
                ai_score = float(sc_mat[j][k])
                rf_name = rapidfuzz_score(q_norm, cat_norms[c_idx])
                rf_catcol = rapidfuzz_score(cust_cat_norm, cat_norms[c_idx]) if cust_cat_norm else 0.0
                rf_best = max(rf_name, rf_catcol)
                # blend: AI embedding is primary signal, rapidfuzz text confirms it
                final = 0.65 * ai_score + 0.35 * rf_best
                if final > best_final:
                    best_final = final
                    best_cat_idx = c_idx

            matched_category[row_i] = cat_raws[best_cat_idx]
            scores[row_i] = round(best_final * 100, 2)

            if best_final >= match_thr:
                methods[row_i] = "ai+rf"
                statuses[row_i] = "MATCHED"
                counters["ai_rf"] += 1
            elif best_final >= review_thr:
                methods[row_i] = "ai+rf"
                statuses[row_i] = "REVIEW"
                counters["review"] += 1
            else:
                methods[row_i] = "ai+rf"
                statuses[row_i] = "UNMATCHED"
                counters["unmatched"] += 1

        processed = batch_start + len(batch_idx)
        elapsed = (dt.datetime.now() - start_ts).total_seconds()
        rate = processed / elapsed if elapsed > 0 else 0
        eta = str(dt.timedelta(seconds=int((total_remaining - processed) / rate))) if rate > 0 else "?"
        log.info(f"  {processed:>7,}/{total_remaining:,} "
                 f"({100 * processed / max(total_remaining,1):5.1f}%) | "
                 f"matched={counters['ai_rf']:,} review={counters['review']:,} "
                 f"unmatched={counters['unmatched']:,} | {rate:.0f}/s ETA={eta}")

    df["matched_lulu_category"] = matched_category
    df["similarity_score"] = scores
    df["match_method"] = methods
    df["match_status"] = statuses
    return df


# --------------------------------------------------------------------------
# OUTPUT
# --------------------------------------------------------------------------
def save_csv(df: pd.DataFrame, out_path: str):
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)

    output_cols = []
    for c in ["material_code", "material_name", "brand", "source_customer",
              "customer_category", "lulu_category"]:
        if c in df.columns:
            output_cols.append(c)

    output_cols += ["matched_lulu_category", "similarity_score", "match_method", "match_status"]

    # Keep the previously-assigned lulu_category (if present, e.g. from a
    # prior LLM pass) alongside the new one for easy comparison.
    if "lulu_category" in df.columns:
        df = df.rename(columns={"lulu_category": "lulu_category_previous"})
        output_cols = [c if c != "lulu_category" else "lulu_category_previous" for c in output_cols]

    out_df = df[output_cols].copy()
    out_df.to_csv(out_path, index=False)
    log.info(f"  CSV saved: {out_path} ({len(out_df):,} rows)")


# --------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------
def main():
    log.info("=" * 70)
    log.info(f"LULU CATEGORY MATCHING PIPELINE v8.0 (CSV-only) — {dt.datetime.now():%Y-%m-%d %H:%M:%S}")
    log.info(f"  Benchmark : {CONFIG['csv_lulu_categories']}")
    log.info(f"  Customer  : {CONFIG['csv_customer_input']}")
    log.info(f"  Output    : {CONFIG['csv_output_path']}")
    log.info(f"  match_threshold={CONFIG['match_threshold']}  review_threshold={CONFIG['review_threshold']}")
    log.info("=" * 70)

    log.info("\n[1/5] Loading LULU category benchmark...")
    cat_raws, cat_norms = load_lulu_categories(CONFIG["csv_lulu_categories"])

    log.info("\n[2/5] Loading customer input...")
    df = load_customer_input(CONFIG["csv_customer_input"])

    log.info("\n[3/5] Loading AI model...")
    model = build_ai_model()

    log.info("\n[4/5] Building category embedding index...")
    cat_index = build_category_index(model, cat_norms)

    log.info("\n[5/5] Matching customer rows to LULU categories...")
    df = match_customers_to_categories(df, cat_raws, cat_norms, model, cat_index)

    save_csv(df, CONFIG["csv_output_path"])

    log.info("\n" + "=" * 70)
    log.info("PIPELINE v8.0 COMPLETE ✓")
    log.info(df["match_status"].value_counts().to_string())
    log.info("=" * 70)


if __name__ == "__main__":
    main()