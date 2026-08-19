"""
Combined Matcher — BARCODE + NAME + WEIGHT + BRAND + HIERARCHY BONUS  (Barcode → Fuzzy → Semantic cascade)
=============================================================================================================
VERSION 5 — CACHED / PARQUET-BACKED  (adapted from V4)  [+ tqdm progress bars]
─────────────────────────────────────────────────────────
SAME MATCHING LOGIC AS V4. This version adds three layers of caching so that
reruns skip work that has already been done:

  1. ★ PARQUET CACHE
     The input Excel file is converted to Parquet exactly once. Every later
     run reads the (much faster) Parquet file instead of re-parsing Excel,
     UNLESS the source Excel file's size/mtime has changed — in which case
     it is re-converted automatically.

  2. ★ SEMANTIC EMBEDDING CACHE (.npy + .json backup)
     Encoding all LULU source names with a local Ollama embedding model
     (SEMANTIC_MODEL, e.g. 'qwen3-embedding:0.6b') is the slowest one-time
     step. The resulting embedding matrix is saved to disk
     as a NumPy .npy file, keyed by a hash of the exact source name list +
     model name. On the next run, if the LULU list and model are unchanged,
     the embeddings are loaded straight from disk instead of recomputed.
     A JSON sidecar (list-of-lists) is also written as a redundant backup —
     if the .npy file is ever corrupted / fails to load, the script falls
     back to rebuilding the array from the .json before giving up and
     recomputing from scratch.

  3. ★ MATCH-RESULT CACHE (.json per customer file)
     The full matched/unmatched output for each customer group is cached
     to its own JSON file, keyed by a hash of that customer's own data
     (descriptions + barcodes) plus the relevant matching config knobs. If
     you rerun the whole script and nothing about that customer's data (or
     the config) changed, that customer's matching is skipped entirely and
     the cached result is reused. If the JSON is corrupted/unreadable, that
     one customer is transparently recomputed (nothing else is affected).

  4. ★ OLLAMA LLM ADJUDICATION (local, no API key/cost)
     Borderline fuzzy/semantic candidates (score in the "grey zone" between
     MIN_CONFIDENCE and ADJUDICATION_HIGH) are sent to a locally-running
     Ollama model for a final yes/no adjudication call, run through a
     thread pool for concurrency. This replaces the GPT-4o-mini adjudication
     step used elsewhere with a free/local model — same idea, no API cost.

★ TQDM PROGRESS BARS (this version)
────────────────────────────────────────────────────
Added live progress bars (via the `tqdm` package) at every long-running
loop in the pipeline, so you can see how far along a run is instead of
watching a wall of print statements. Specifically:
  - Outer loop over customer files (main())
  - Per-customer fuzzy-matching group loop (match_customer())
  - Per-customer semantic-rescue loop (match_customer())
  - Ollama grey-zone adjudication waves (ollama_adjudicate_batch())
  - Ollama embedding chunk requests (ollama_embed_texts())
Any print() that used to fire *inside* one of these loops was changed to
tqdm.write() instead — a plain print() while a bar is active corrupts the
bar's rendering (it gets pushed around/duplicated on screen), whereas
tqdm.write() prints a clean line above the bar and lets the bar keep
redrawing itself in place below it.

FOLDER LAYOUT (created automatically if missing)
────────────────────────────────────────────────
  C:\\Users\\HP\\Desktop\\zero1\\            <- project root (venv lives here)
  C:\\Users\\HP\\Desktop\\zero1\\data\\      <- input files + parquet cache
  C:\\Users\\HP\\Desktop\\zero1\\cache\\     <- embedding + match-result cache
  C:\\Users\\HP\\Desktop\\zero1\\output\\    <- final Excel output

Put your source Excel file in the `data\\` folder (or point INPUT_FILE at
wherever it currently is — it will be copied/converted into `data\\` on
first run).
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
OUTPUT_FILE  = OUTPUT_DIR + r"\ta.xlsx"

# ★ CACHE — set True to ignore every cache (parquet / embeddings / match
# results) and redo the entire pipeline from scratch. Use this after you
# change the matching LOGIC itself (not just the data), since the cache
# keys don't know about code changes.
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
# customer's barcode "learned" against the matched LULU row.
BACKFILL_MIN_CONFIDENCE = 89

MIN_CONFIDENCE     = 75
ENSEMBLE_FUZZY     = True
USE_SEMANTIC       = True

# ★ OLLAMA EMBEDDINGS — semantic layer now runs on a local Ollama embedding
# model instead of sentence-transformers. No GPU needed, no separate model
# download outside Ollama, everything stays local. Pick a size that fits
# your RAM:
#   qwen3-embedding:0.6b  <- recommended for laptops / low-RAM (no GPU)
#   qwen3-embedding:4b    <- needs more RAM, still CPU-runnable
#   qwen3-embedding:8b    <- heavy; only if you have a GPU or lots of free RAM
SEMANTIC_MODEL     = "qwen3-embedding:8b"   # this is now an OLLAMA model name (`ollama pull qwen3-embedding:0.6b`)
SEMANTIC_BATCH     = 64     # texts sent per Ollama /api/embed call
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

# ── ★ OLLAMA LLM ADJUDICATION ──────────────────────────────────────────────
# Local, free, no API key. Requires `ollama serve` running and the model
# already pulled, e.g.:   ollama pull llama3.1
USE_OLLAMA_ADJUDICATION = True
OLLAMA_HOST             = "http://localhost:11434"
OLLAMA_CHAT_URL         = OLLAMA_HOST + "/api/chat"
OLLAMA_MODEL            = "llama3.1"     # swap for whatever local model you have pulled
OLLAMA_TIMEOUT_SECS     = 30
OLLAMA_MAX_RETRIES      = 2

# Any fuzzy/semantic candidate whose score falls in [MIN_CONFIDENCE, this)
# is considered a "grey zone" match and gets sent to the LLM for a final
# strict yes/no adjudication instead of being auto-accepted on score alone.
ADJUDICATION_HIGH_CONFIDENCE = 90

# ★ WORKER POOL — concurrent Ollama calls. Ollama serves requests one at a
# time per model by default on modest hardware, so this is tuned modestly;
# raise it if you have a beefier GPU/host or are running multiple model
# instances behind a load balancer.
OLLAMA_WORKERS   = 12   # thread-pool size for concurrent adjudication calls
OLLAMA_POOL_SIZE = 6    # number of pairs sent to the pool per "wave" (batching)

# ── ★ OLLAMA EMBEDDING ENDPOINT (semantic layer) ───────────────────────────
OLLAMA_EMBED_URL           = OLLAMA_HOST + "/api/embed"
OLLAMA_EMBED_TIMEOUT_SECS  = 60
OLLAMA_EMBED_MAX_RETRIES   = 2
OLLAMA_EMBED_WORKERS       = 4   # concurrent chunk requests for embedding calls

# ===========================================================================

import re, sys, time, warnings, hashlib, json
import os as _os
import concurrent.futures
import numpy as np
import pandas as pd
from collections import defaultdict
from pathlib import Path
from rapidfuzz import fuzz, process
from tqdm import tqdm  # ★ TQDM — progress bars for long-running loops

try:
    import requests
except ImportError:
    requests = None

warnings.filterwarnings("ignore")

# ── ★ CACHE / FOLDER SETUP ──────────────────────────────────────────────────
for _d in (PROJECT_ROOT, DATA_DIR, CACHE_DIR, OUTPUT_DIR):
    _os.makedirs(_d, exist_ok=True)


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


# ===========================================================================
# ★ OLLAMA — strict adjudication layer
# ===========================================================================

OLLAMA_SYSTEM_PROMPT = """You are a product-matching adjudicator for a Qatar retail reconciliation system.
Decide if a CUSTOMER PRODUCT (AL_MEERA/C4/TALABAT/GRAND_MALL) is the SAME
real-world product as a CANDIDATE from the LULU master catalogue.

MATCH = same brand (or known alias) + same product + same variant/flavor +
same pack size (unit-format differences OK: 1L=1000ML, 0.5KG=500G).

HARD REJECT (no match, even at 95%+ text similarity) if these conflict:
- oil/fat type (sunflower ≠ olive ≠ corn ≠ canola)
- flavor/variant (chocolate chip ≠ oatmeal, classic ≠ pizza ≠ smoked)
- pack size/count (real unit mismatch, not notation)
- product form (frozen ≠ fresh, powder ≠ liquid, whole ≠ ground)
- brand family, unless it's a known alias (CADBURYS IMP→CADBURY, AL KABEER→
  KABEER, NESTLE MAGGI→MAGGI) or a sourcing prefix (AUS BEEF/NZ BEEF→BEEF)

IGNORE (not conflicts): word order, abbreviations (ORG=ORGANIC), marketing
filler (NEW/PREMIUM/FRESH), transliteration variants.

Text similarity never overrides an attribute conflict.

EXAMPLES:

Input:
CUSTOMER: "CADBURYS IMP DAIRY MILK 45G" | source: TALABAT
CANDIDATE: [L212] "CADBURY DAIRY MILK CHOCOLATE 45G" | 0.88 sim
Output:
{"match": true, "matched_lulu_code": "L212", "confidence": 88,
 "conflict_flags": [], "reasoning": "CADBURYS IMP is CADBURY's alias; product/size match."}

Input:
CUSTOMER: "AL AIN SUNFLOWER OIL 1.5L" | source: C4
CANDIDATE: [L045] "AL AIN CORN OIL 1.5L" | 0.94 sim
Output:
{"match": false, "matched_lulu_code": null, "confidence": 97,
 "conflict_flags": ["oil_type"], "reasoning": "Same brand/size, but oil type conflicts."}

Input:
CUSTOMER: "AL KABEER CHICKEN NUGGETS 0.5KG" | source: AL_MEERA
CANDIDATE: [L177] "KABEER CHICKEN NUGGETS 500G" | 0.90 sim
Output:
{"match": true, "matched_lulu_code": "L177", "confidence": 93,
 "conflict_flags": [], "reasoning": "0.5KG=500G; KABEER is standardized short form of AL KABEER."}

Input:
CUSTOMER: "PARLE MONACO SMOKED CRACKERS 200G" | source: GRAND_MALL
CANDIDATE: [L309] "PARLE MONACO CLASSIC CRACKERS 200G" | 0.86 sim
Output:
{"match": false, "matched_lulu_code": null, "confidence": 78,
 "conflict_flags":
"""


def _build_ollama_user_prompt(src_desc, src_brand, src_weight, cust_desc, cust_brand, cust_weight):
    def _fmt_w(w):
        return f"{w:.1f} (normalized g/ml)" if w is not None and not (isinstance(w, float) and pd.isna(w)) else "unknown"
    return (
        f"SOURCE (LULU):\n"
        f"  Name: {src_desc}\n"
        f"  Brand: {src_brand if src_brand else 'unknown'}\n"
        f"  Weight/Size: {_fmt_w(src_weight)}\n\n"
        f"CANDIDATE (customer):\n"
        f"  Name: {cust_desc}\n"
        f"  Brand: {cust_brand if cust_brand else 'unknown'}\n"
        f"  Weight/Size: {_fmt_w(cust_weight)}\n\n"
        f"Are these the same product? Reply with the JSON object only."
    )


def _parse_ollama_json(raw_text: str):
    """Best-effort extraction of the {match, confidence, reason} JSON blob
    from a model response that might include stray text around it."""
    if not raw_text:
        return None
    text = raw_text.strip()
    try:
        return json.loads(text)
    except Exception:
        pass
    m = re.search(r'\{.*\}', text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except Exception:
            return None
    return None


def ollama_adjudicate_pair(src_desc, src_brand, src_weight, cust_desc, cust_brand, cust_weight):
    """
    Single synchronous call to the local Ollama chat endpoint for one
    candidate pair. Returns a dict {'match': bool, 'confidence': int,
    'reason': str} — on any failure (Ollama not running, timeout, bad
    JSON, etc.) returns {'match': None, 'confidence': 0, 'reason': 'error: ...'}
    so callers can fall back to the fuzzy/semantic score instead.
    """
    if requests is None:
        return {'match': None, 'confidence': 0, 'reason': 'error: requests library not installed'}

    payload = {
        "model": OLLAMA_MODEL,
        "messages": [
            {"role": "system", "content": OLLAMA_SYSTEM_PROMPT},
            {"role": "user", "content": _build_ollama_user_prompt(
                src_desc, src_brand, src_weight, cust_desc, cust_brand, cust_weight)},
        ],
        "stream": False,
        "options": {"temperature": 0},
    }

    last_err = None
    for attempt in range(OLLAMA_MAX_RETRIES + 1):
        try:
            resp = requests.post(OLLAMA_CHAT_URL, json=payload, timeout=OLLAMA_TIMEOUT_SECS)
            resp.raise_for_status()
            data = resp.json()
            content = (data.get("message") or {}).get("content", "")
            parsed = _parse_ollama_json(content)
            if parsed is None:
                last_err = f"unparseable response: {content[:200]!r}"
                continue
            return {
                'match':      bool(parsed.get('match', False)),
                'confidence': int(parsed.get('confidence', 0) or 0),
                'reason':     str(parsed.get('reason', ''))[:200],
            }
        except Exception as e:
            last_err = str(e)
            time.sleep(0.5 * (attempt + 1))

    return {'match': None, 'confidence': 0, 'reason': f'error: {last_err}'}


def ollama_adjudicate_batch(pairs: list, pool_size: int = OLLAMA_POOL_SIZE,
                             max_workers: int = OLLAMA_WORKERS) -> list:
    """
    Runs ollama_adjudicate_pair() over a list of pair-dicts concurrently
    using a ThreadPoolExecutor(max_workers=OLLAMA_WORKERS), processed in
    waves of `pool_size` at a time so progress can be logged and memory
    stays bounded on very large grey-zone queues. Returns results in the
    same order as `pairs`.

    Each item in `pairs` must be a tuple:
        (src_desc, src_brand, src_weight, cust_desc, cust_brand, cust_weight)
    """
    results = [None] * len(pairs)
    if not pairs:
        return results

    n_waves = (len(pairs) + pool_size - 1) // pool_size
    # ★ TQDM — one tick per wave of adjudication calls
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        for w in tqdm(range(n_waves), desc="  Ollama adjudication", unit="wave"):
            start = w * pool_size
            end   = min(start + pool_size, len(pairs))
            chunk_indices = list(range(start, end))

            future_to_idx = {
                executor.submit(ollama_adjudicate_pair, *pairs[i]): i
                for i in chunk_indices
            }
            for fut in concurrent.futures.as_completed(future_to_idx):
                idx = future_to_idx[fut]
                try:
                    results[idx] = fut.result()
                except Exception as e:
                    results[idx] = {'match': None, 'confidence': 0, 'reason': f'error: {e}'}

            if (w + 1) % 5 == 0 or (w + 1) == n_waves:
                # ★ TQDM — tqdm.write() instead of print() so this line doesn't
                # break/duplicate the live progress bar rendering above it.
                tqdm.write(f"    Ollama adjudication wave {w+1}/{n_waves}  "
                           f"({end}/{len(pairs)} pairs)")

    return results


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
    if not s or set(s) == {'0'} or len(s) < 6:
        return None
    return s


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
# ★ SEMANTIC LAYER — Ollama embedding model, with .npy + .json backup cache
# ===========================================================================

def _ollama_embed_chunk(chunk: list, model: str) -> list:
    """
    One /api/embed call for a chunk of texts. Returns a list of raw
    embedding vectors (lists of floats), in the same order as `chunk`.
    Retries a few times before raising, so callers running this inside a
    thread pool get a clean exception instead of a silent None.
    """
    if requests is None:
        raise ImportError("Run:  pip install requests --break-system-packages")

    payload = {"model": model, "input": chunk}
    last_err = None
    for attempt in range(OLLAMA_EMBED_MAX_RETRIES + 1):
        try:
            resp = requests.post(OLLAMA_EMBED_URL, json=payload, timeout=OLLAMA_EMBED_TIMEOUT_SECS)
            resp.raise_for_status()
            data = resp.json()
            embs = data.get("embeddings")
            if not embs or len(embs) != len(chunk):
                raise ValueError(f"unexpected /api/embed response shape "
                                  f"(got {0 if not embs else len(embs)}, expected {len(chunk)})")
            return embs
        except Exception as e:
            last_err = str(e)
            time.sleep(0.5 * (attempt + 1))

    raise RuntimeError(
        f"Ollama embedding call failed after {OLLAMA_EMBED_MAX_RETRIES + 1} attempts: {last_err}\n"
        f"  → Check that 'ollama serve' is running and the model is pulled:\n"
        f"    ollama pull {model}"
    )


def ollama_embed_texts(texts: list, model: str = SEMANTIC_MODEL,
                        batch_size: int = SEMANTIC_BATCH,
                        max_workers: int = OLLAMA_EMBED_WORKERS,
                        label: str = "") -> np.ndarray:
    """
    Embeds a list of texts using a local Ollama embedding model, sending
    `batch_size` texts per HTTP call, with up to `max_workers` chunk
    requests in flight at once via a ThreadPoolExecutor. Returns an
    (N, dim) float32 array, L2-normalized row-wise so a plain dot product
    gives cosine similarity (matches how the rest of the pipeline scores
    semantic candidates).
    """
    if not texts:
        return np.zeros((0, 0), dtype=np.float32)

    n_chunks = (len(texts) + batch_size - 1) // batch_size
    chunks = [texts[i * batch_size:(i + 1) * batch_size] for i in range(n_chunks)]
    results = [None] * n_chunks

    tag = f"[{label}] " if label else ""
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_idx = {
            executor.submit(_ollama_embed_chunk, chunk, model): i
            for i, chunk in enumerate(chunks)
        }
        # ★ TQDM — one tick per completed embedding chunk request
        for fut in tqdm(concurrent.futures.as_completed(future_to_idx),
                         total=n_chunks, desc=f"  {tag}Embedding", unit="chunk"):
            idx = future_to_idx[fut]
            results[idx] = fut.result()   # raises if that chunk ultimately failed

    flat = [vec for chunk_res in results for vec in chunk_res]
    arr = np.array(flat, dtype=np.float32)
    norms = np.linalg.norm(arr, axis=1, keepdims=True)
    norms[norms == 0] = 1.0
    return arr / norms


class SemanticMatcher:
    """
    Semantic candidate-rescue layer backed by a local Ollama embedding
    model (SEMANTIC_MODEL, e.g. 'qwen3-embedding:0.6b') instead of
    sentence-transformers. Same public interface as before
    (encode_batch / scores_for_query) so match_customer() doesn't need to
    change at all.
    """

    def __init__(self, source_names: list):
        cache_key = _hash_list(source_names) + '_' + SEMANTIC_MODEL.replace('/', '_').replace(':', '_')
        npy_path  = _os.path.join(CACHE_DIR, f"src_emb_ollama_{cache_key}.npy")
        json_path = _os.path.join(CACHE_DIR, f"src_emb_ollama_{cache_key}.json")

        print(f"  Semantic layer: Ollama embedding model '{SEMANTIC_MODEL}' @ {OLLAMA_EMBED_URL}")
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
                    try:
                        np.save(npy_path, self.src_emb)
                    except Exception:
                        pass
                else:
                    print("  ⚠ .json cache row count mismatch — recomputing")
            except Exception as e:
                print(f"  ⚠ .json cache corrupted ({e}) — recomputing")

        # ── both caches missing/corrupt → recompute via Ollama ──────────
        if self.src_emb is None:
            print(f"  Encoding {len(source_names):,} source names via Ollama "
                  f"(model={SEMANTIC_MODEL}, batch={SEMANTIC_BATCH}, workers={OLLAMA_EMBED_WORKERS}) …")
            self.src_emb = ollama_embed_texts(
                source_names, model=SEMANTIC_MODEL, batch_size=SEMANTIC_BATCH,
                max_workers=OLLAMA_EMBED_WORKERS, label="LULU source",
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
        return ollama_embed_texts(
            names, model=SEMANTIC_MODEL, batch_size=SEMANTIC_BATCH,
            max_workers=OLLAMA_EMBED_WORKERS,
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
    src_material_code_display,                     # ★ MATERIAL CODE
    name_key_to_src,                               # ★ BRIDGE
):
    t0 = time.time()

    print(f"\n{'='*60}")
    print(f"  Customer: {label}")
    print(f"{'='*60}")

    cust['source_file'] = label
    print(f"  Rows: {len(cust):,}")

    def _maybe_backfill(cust_barcode_raw, src_i, method, conf_int):
        cust_bc_norm = normalize_barcode(cust_barcode_raw)
        if not cust_bc_norm:
            return False, None, False
        is_high_conf = method in ('Fuzzy', 'Semantic') and conf_int is not None and conf_int >= BACKFILL_MIN_CONFIDENCE
        is_bridge    = method == 'Barcode (Bridged)'
        if not (is_high_conf or is_bridge):
            return False, None, False
        newly_learned = cust_bc_norm not in barcode_to_src
        barcode_to_src[cust_bc_norm] = src_i
        if not src_barcode_display[src_i]:
            src_barcode_display[src_i] = cust_barcode_raw
        return True, cust_barcode_raw, newly_learned

    cust['_barcode'] = cust['barcode'].apply(normalize_barcode)
    barcode_matched_cis = {}   # ci -> (src_i, method_label)

    if barcode_to_src:
        for ci, bc in enumerate(cust['_barcode'].tolist()):
            if bc and bc in barcode_to_src:
                barcode_matched_cis[ci] = (barcode_to_src[bc], 'Barcode')

    n_bridged = 0
    if name_key_to_src:
        for ci in range(len(cust)):
            if ci in barcode_matched_cis:
                continue
            key = normalize_name_key(cust.at[ci, 'desc'])
            if key and key in name_key_to_src:
                barcode_matched_cis[ci] = (name_key_to_src[key], 'Barcode (Bridged)')
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
    # ★ TQDM — one tick per (kind, weight-band) group of customer rows being
    # fuzzy-matched against the LULU catalogue for this customer file
    for g_idx, ((kind, band), cidxs) in enumerate(
        tqdm(cust_groups.items(), total=total_groups,
             desc=f"  [{label}] fuzzy groups", unit="grp", leave=False), 1
    ):

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
            # ★ TQDM — tqdm.write() keeps this status line from corrupting
            # the live "fuzzy groups" bar above it
            tqdm.write(f"  fuzzy group {g_idx}/{total_groups}  ({time.time()-t0:.1f}s)")

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
        # ★ TQDM — one tick per customer row being rescued via the
        # semantic (embedding) layer after failing/scoring low on fuzzy
        for count, ci in enumerate(
            tqdm(needs_semantic, desc=f"  [{label}] semantic rescue", unit="row", leave=False), 1
        ):
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
                tqdm.write(f"  semantic {count}/{len(needs_semantic)}  ({time.time()-t0:.1f}s)")

    # ★ OLLAMA — grey-zone adjudication pass. Any row whose fuzzy/semantic
    # result landed as 'ok' or 'ok_no_brand' with a score between
    # MIN_CONFIDENCE and ADJUDICATION_HIGH_CONFIDENCE is not auto-accepted;
    # it's sent to the local Ollama model (via a thread pool) for a strict
    # brand + weight + name-variant check before being finalized.
    if USE_OLLAMA_ADJUDICATION:
        grey_cis = [
            ci for ci, r in enumerate(final_results)
            if r is not None and r[2] in ('ok', 'ok_no_brand')
            and MIN_CONFIDENCE <= r[1] < ADJUDICATION_HIGH_CONFIDENCE
        ]
        if grey_cis:
            print(f"  Ollama adjudication queue: {len(grey_cis):,} grey-zone row(s) "
                  f"(model={OLLAMA_MODEL}, pool={OLLAMA_POOL_SIZE}, workers={OLLAMA_WORKERS})")
            pairs = []
            for ci in grey_cis:
                si = final_results[ci][0]
                pairs.append((
                    src_name[si], src_brand_display[si], src_weight[si],
                    cust_name[ci], cust.at[ci, 'brand'], cust_weight[ci],
                ))
            verdicts = ollama_adjudicate_batch(pairs, pool_size=OLLAMA_POOL_SIZE, max_workers=OLLAMA_WORKERS)

            n_confirmed = n_rejected = n_unavailable = 0
            for ci, verdict in zip(grey_cis, verdicts):
                si, sc, reason, method = final_results[ci]
                if verdict['match'] is None:
                    # Ollama unreachable/errored — keep the original fuzzy/semantic
                    # verdict untouched rather than silently dropping the row.
                    n_unavailable += 1
                    continue
                if verdict['match']:
                    new_sc = max(sc, float(verdict['confidence']))
                    final_results[ci] = (si, new_sc, reason, f'{method}+Ollama')
                    n_confirmed += 1
                else:
                    final_results[ci] = (None, float(verdict['confidence']), 'ollama_rejected', method)
                    n_rejected += 1

            print(f"  Ollama verdicts: {n_confirmed:,} confirmed | {n_rejected:,} rejected "
                  f"| {n_unavailable:,} unavailable (fell back to prior score)")

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
            src_i, bc_method = barcode_matched_cis[ci]

            borrowed, borrowed_bc, newly_learned = _maybe_backfill(cust_barcode_raw, src_i, bc_method, None)
            if newly_learned:
                n_backfilled += 1

            src_wt = src_weight[src_i]
            src_pk = src_pack[src_i]
            src_brand_raw = src_brand_display[src_i]
            blabel = brand_label(src_brand_raw, cust_brand_raw, BRAND_FUZZY_THRESH, alias_dict)
            is_bridged = (bc_method == 'Barcode (Bridged)')

            matched_src_indices.add(src_i)
            matched_rows.append({
                'Source File':                     label,
                'Our Brand (brand_standardized)':  src_brand_raw,
                'Our Name (parsed)':                src_name[src_i],
                'Our Material Code':                src_material_code_display[src_i],
                'Our Weight g/ml':                  src_wt,
                'Our Pack':                          src_pk,
                'Our Description (material_name)':  src_descraw[src_i],
                'Our Barcode':                       src_barcode_display[src_i],
                'Borrowed Barcode':                  borrowed_bc,
                'Customer Description':             cust_desc_raw,
                'Customer Brand':                   cust_brand_raw,
                'Customer Material Code':           cust_material_code_raw,
                'Customer Barcode':                 cust_barcode_raw,
                'Customer Name (parsed)':           cust_name[ci],
                'Customer Weight g/ml':             cust_wt,
                'Customer Pack':                     cust_pk,
                'Match Status':                     'Matched (Barcode-Bridged)' if is_bridged else 'Matched (Barcode)',
                'Confidence Score':                 97 if is_bridged else 100,
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

        if reason in ('brand_not_in_catalogue', 'brand_mismatch', 'ollama_rejected') or src_i is None:
            unmatched_rows.append({**base, 'Reason': reason, 'Best Score': int(round(conf))})
            continue

        conf_int = int(round(conf))
        if conf_int < MIN_CONFIDENCE:
            unmatched_rows.append({**base, 'Reason': 'below_threshold', 'Best Score': conf_int})
            continue

        borrowed, borrowed_bc, newly_learned = _maybe_backfill(cust_barcode_raw, src_i, method, conf_int)
        if newly_learned:
            n_backfilled += 1

        src_wt        = src_weight[src_i]
        src_pk        = src_pack[src_i]
        src_brand_raw = src_brand_display[src_i]
        blabel        = brand_label(src_brand_raw, cust_brand_raw, BRAND_FUZZY_THRESH, alias_dict)
        status        = 'Matched (High)' if conf_int >= 85 else 'Matched (Medium)'

        matched_src_indices.add(src_i)

        matched_rows.append({
            'Source File':                  label,
            'Our Brand (brand_standardized)': src_brand_raw,
            'Our Name (parsed)':             src_name[src_i],
            'Our Material Code':             src_material_code_display[src_i],
            'Our Weight g/ml':               src_wt,
            'Our Pack':                      src_pk,
            'Our Description (material_name)': src_descraw[src_i],
            'Our Barcode':                    src_barcode_display[src_i],
            'Borrowed Barcode':               borrowed_bc,
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

    return matched_rows, unmatched_rows, alias_dict, matched_src_indices


# ===========================================================================
# ★ MATCH-RESULT CACHE — per-customer JSON, with corruption fallback
# ===========================================================================

def _customer_cache_path(label, cust, config_sig) -> str:
    cache_input = cust['desc'].astype(str).tolist() + cust['barcode'].astype(str).tolist()
    key = _hash_list(cache_input) + '_' + config_sig
    safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', str(label))[:60]
    return _os.path.join(CACHE_DIR, f"match_{safe_label}_{key}.json")


def load_customer_cache(path, barcode_to_src, src_barcode_display):
    """Returns (m_rows, u_rows, aliases, src_idx_set) or None if cache is
    missing/corrupt/unusable. Also replays any barcodes this customer
    taught the catalogue last time, so downstream customers still benefit."""
    if not _os.path.exists(path):
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            cached = json.load(f)
        m_rows      = cached['matched']
        u_rows      = cached['unmatched']
        aliases     = cached['aliases']
        src_idx_set = set(cached['src_idx'])
        for bc, si in cached.get('learned_barcodes', {}).items():
            barcode_to_src[bc] = si
            if not src_barcode_display[si]:
                src_barcode_display[si] = cached.get('learned_barcode_display', {}).get(bc)
        return m_rows, u_rows, aliases, src_idx_set
    except Exception as e:
        print(f"  ⚠ Match-result cache corrupted/unreadable ({e}) — recomputing this customer")
        return None


def save_customer_cache(path, m_rows, u_rows, aliases, src_idx_set, learned, learned_display):
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(_json_safe({
                'matched':                m_rows,
                'unmatched':               u_rows,
                'aliases':                 aliases,
                'src_idx':                 list(src_idx_set),
                'learned_barcodes':        learned,
                'learned_barcode_display': learned_display,
            }), f)
        print(f"  ✓ Cached match results → {path}")
    except Exception as e:
        print(f"  ⚠ Could not write match-result cache: {e}")


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
            fuzzy = sum(1 for r in m_rows if r['Method'].startswith('Fuzzy'))
            sem   = sum(1 for r in m_rows if r['Method'].startswith('Semantic'))
            ollama_confirmed = sum(1 for r in m_rows if 'Ollama' in r['Method'])
            cov_n = len(src_idx)
            cov_p = f"{cov_n / total_src_rows * 100:.1f}%" if total_src_rows else '0%'

            rows += [
                {'Metric': f'[{label}] Customer rows',              'Value': t,           'Detail': ''},
                {'Metric': f'[{label}] Matched',                    'Value': len(m_rows), 'Detail': match_pct},
                {'Metric': f'[{label}] Unmatched',                  'Value': len(u_rows), 'Detail': unmatch_pct},
                {'Metric': f'[{label}] via Barcode',                'Value': barcode,     'Detail': ''},
                {'Metric': f'[{label}] via Barcode (Bridged)',      'Value': bridged,     'Detail': ''},
                {'Metric': f'[{label}] High confidence (≥85)',      'Value': high,        'Detail': f"{high/len(m_rows)*100:.1f}% of matched" if m_rows else ''},
                {'Metric': f'[{label}] Med confidence (75-84)',     'Value': med,         'Detail': f"{med/len(m_rows)*100:.1f}% of matched" if m_rows else ''},
                {'Metric': f'[{label}] via Fuzzy',                  'Value': fuzzy,       'Detail': ''},
                {'Metric': f'[{label}] via Semantic',               'Value': sem,         'Detail': ''},
                {'Metric': f'[{label}] confirmed by Ollama',        'Value': ollama_confirmed, 'Detail': 'grey-zone rows confirmed by local LLM'},
                {'Metric': f'[{label}] Source SKUs covered',        'Value': cov_n,       'Detail': f"{cov_p} of our {total_src_rows:,} source SKUs"},
                {'Metric': '', 'Value': '', 'Detail': ''},
            ]

        rows.append({'Metric': '─── COMBINED TOTALS (ALL FILES) ───', 'Value': '', 'Detail': ''})
        rows.append({'Metric': '', 'Value': '', 'Detail': ''})
        rows += [
            {'Metric': 'Total customer rows',    'Value': total_cust_rows,    'Detail': ''},
            {'Metric': 'Total matched',          'Value': len(all_matched),   'Detail': f"{len(all_matched)/total_cust_rows*100:.1f}% of all customer rows" if total_cust_rows else ''},
            {'Metric': 'Total unmatched',        'Value': len(all_unmatched), 'Detail': f"{len(all_unmatched)/total_cust_rows*100:.1f}% of all customer rows" if total_cust_rows else ''},
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

    print(f"  Total matched  : {len(matched_df):,}")
    print(f"  Total unmatched: {len(unmatched_df):,}")
    print(f"  Source coverage: {src_covered_count:,} / {total_src_rows:,} ({src_covered_pct:.1f}%)")


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
    if USE_SEMANTIC:
        print(f"  Semantic embeddings ON — model='{SEMANTIC_MODEL}' host={OLLAMA_HOST} "
              f"batch={SEMANTIC_BATCH} workers={OLLAMA_EMBED_WORKERS}")
    if USE_OLLAMA_ADJUDICATION:
        print(f"  Ollama adjudication ON — model='{OLLAMA_MODEL}' host={OLLAMA_HOST} "
              f"pool_size={OLLAMA_POOL_SIZE} workers={OLLAMA_WORKERS}")
    if (USE_SEMANTIC or USE_OLLAMA_ADJUDICATION) and requests is None:
        print("  ⚠ 'requests' package not installed — Ollama calls (embeddings + adjudication) "
              "will fail. Run: pip install requests --break-system-packages")

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

    name_key_to_src = {}
    if BARCODE_COL and BARCODE_COL in df.columns:
        name_to_barcodes = defaultdict(set)
        for nm, bc_raw in zip(df[NAME_COL].tolist(), df[BARCODE_COL].tolist()):
            key = normalize_name_key(nm)
            bc  = normalize_barcode(bc_raw)
            if key and bc:
                name_to_barcodes[key].add(bc)

        bridge_conflicts = 0
        for key, bcs in name_to_barcodes.items():
            hits = {barcode_to_src[bc] for bc in bcs if bc in barcode_to_src}
            if not hits:
                continue
            if len(hits) > 1:
                bridge_conflicts += 1
                continue
            name_key_to_src[key] = next(iter(hits))

        print(f"  Barcode↔material_name bridge: {len(name_key_to_src):,} material name(s) "
              f"resolvable via a co-occurring barcode elsewhere in the file"
              + (f"  ({bridge_conflicts:,} ambiguous — multiple different LULU rows implicated, skipped)"
                 if bridge_conflicts else ""))
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

    # signature of the config knobs that affect match_customer's output,
    # so a config change (not just data change) invalidates old caches
    config_sig = _hash_list([
        MIN_CONFIDENCE, BRAND_FUZZY_THRESH, WEIGHT_TOL, BACKFILL_MIN_CONFIDENCE,
        REQUIRE_BRAND_MATCH, ALLOW_MATCH_WITHOUT_BRAND, DEDUPE_CUSTOMER_BY_DESC,
        SEMANTIC_MODEL, ENSEMBLE_FUZZY, USE_SEMANTIC,
        USE_OLLAMA_ADJUDICATION, OLLAMA_MODEL, ADJUDICATION_HIGH_CONFIDENCE,
    ])

    all_matched             = []
    all_unmatched           = []
    all_aliases             = {}
    per_file                = {}
    all_matched_src_indices = set()

    # ★ TQDM — outer, top-level bar: one tick per customer file processed
    # (AL_MEERA, C4, TALABAT, GRAND_MALL, etc.) — this is the bar that best
    # reflects "how much of the whole run is done".
    for label in tqdm(customer_labels, desc="Customer files", unit="file"):
        cust = get_customer_frame(df, label, barcode_to_src=barcode_to_src)
        if cust.empty:
            continue

        # ★ CACHE — check per-customer JSON cache first
        cache_path = _customer_cache_path(label, cust, config_sig)
        cached = None
        if not FORCE_RECOMPUTE:
            cached = load_customer_cache(cache_path, barcode_to_src, src_barcode_display)

        if cached is not None:
            tqdm.write(f"\n  ✓ [{label}] Match-result cache hit — skipping recompute → {cache_path}")
            m_rows, u_rows, aliases, src_idx_set = cached
        else:
            barcodes_before = dict(barcode_to_src)   # snapshot to detect what THIS run learned

            m_rows, u_rows, aliases, src_idx_set = match_customer(
                label, cust,
                src_name, src_weight, src_kind, src_pack, src_brands,
                src_descraw, src_brand_display, src_buckets, all_src_idxs,
                known_src_brands, semantic,
                src_h1, src_h2,
                barcode_to_src, src_barcode_display,
                src_material_code_display,
                name_key_to_src,
            )

            learned         = {bc: si for bc, si in barcode_to_src.items() if bc not in barcodes_before}
            learned_display = {bc: src_barcode_display[si] for bc, si in learned.items()}

            save_customer_cache(cache_path, m_rows, u_rows, aliases, src_idx_set, learned, learned_display)

        all_matched.extend(m_rows)
        all_unmatched.extend(u_rows)
        all_aliases.update(aliases)
        per_file[label] = (m_rows, u_rows, src_idx_set)
        all_matched_src_indices |= src_idx_set

    write_all_results(
        all_matched, all_unmatched, all_aliases, per_file,
        total_src_rows, all_matched_src_indices, src_descraw,
    )

    print(f"\nTotal time: {time.time()-t_total:.1f}s")
    print(f"Saved → {OUTPUT_FILE}")


if __name__ == '__main__':
    main()