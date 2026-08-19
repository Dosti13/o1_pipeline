"""
QNIE v9 — COST-OPTIMISED (cache-first, shorter prompts, row dedup)
+ RETRY / EXPONENTIAL BACKOFF PATCH
+ ★ THREADPOOL PATCH (4 concurrent LLM workers, 12-connection HTTP pool)
+ ★★ BRAND/WEIGHT INTEGRITY PATCH (this version)
===================================================================
Changes vs v8:
  1. ROW-LEVEL MATCH CACHE  — every resolved row (barcode, embedding,
     or LLM) is written to a single persistent JSON file keyed by a
     hash of (description, brand, barcode). Next run skips it entirely.
  2. QUERY DEDUPLICATION    — identical parsed names are embedded once;
     the result is reused for all copies, cutting embedding tokens.
  3. SHORTER LLM PROMPT     — system prompt cut from ~1200→~350 tokens,
     user prompt compacted. Saves ~800 input tokens PER LLM call.
  4. EMBEDDING CACHE FOR QUERIES — customer query embeddings are cached
     to .npy so re-runs with the same customer data cost zero tokens.

Changes in retry / backoff patch:
  5. Custom `call_with_backoff()` wraps every OpenAI API call (both
     embeddings.create and chat.completions.create) with:
       - configurable max attempts
       - exponential backoff (base_delay * 2**attempt)
       - random jitter to avoid thundering-herd retries
       - full exception logging per attempt (so a hang/failure is
         VISIBLE instead of silently freezing the whole run)
  6. Per-batch progress printing (every batch, not every 10th) with
     flush=True, so Windows buffered stdout doesn't hide where a
     script is stuck.
  7. OpenAI client instantiated with a request-level `timeout` and
     `max_retries=0` (SDK auto-retry disabled) because we handle
     retries ourselves with visible logging + backoff control.

Changes in THREADPOOL patch (concurrency):
  8. LLM adjudication calls now run concurrently via a
     ThreadPoolExecutor (LLM_MAX_WORKERS, default 4) instead of one
     call at a time.
  9. The OpenAI client uses an httpx.Client with a larger connection
     pool (HTTP_POOL_SIZE, default 12).
  10. RowMatchCache is thread-safe (guarded by a re-entrant lock).

★★ Changes in THIS patch (brand/weight integrity):
  11. LLM SYSTEM PROMPT now states explicit HARD RULES: a candidate is
      NOT a match if weight/pack differs OR if brand differs (brand
      identity outranks category/description similarity). Previously
      the prompt only enforced weight, leaving brand entirely to the
      model's discretion — this caused wrong-brand matches like
      "SEARA" -> "Asaffa" and "AL WAHA" -> "Waha" chicken products.
  12. POST-LLM HARD VETO — the LLM's chosen candidate is no longer
      trusted blindly. Each candidate's brand_ok/weight_ok flags
      (computed the same way as the embedding auto-accept path, via
      brands_overlap()/weights_match()) are carried through to the
      thread-pool result handler. If the LLM picks a candidate that
      fails either check, the result is force-rejected
      (`hard_reject_brand_conflict` / `hard_reject_weight_conflict`)
      regardless of the LLM's stated confidence, and NOTHING gets
      written to the row cache as a false match.
  13. VALIDATED BARCODE BACKFILL — `_maybe_backfill()` previously let
      ANY high-confidence Embedding/LLM match "teach" a new
      barcode -> source-row link, with no check that the match was
      actually plausible. A single bad high-confidence match (e.g.
      "LESTELLO CHICKPEA CAKES 130GM" incorrectly linked to
      "ULKER OLALA SUFLE CAKE 70g") would then get silently inherited
      as a "Matched (Barcode), 100% confidence" result by every other
      customer row sharing that barcode, forever, with no further
      brand/weight checking. Backfill now requires weights_match() to
      not be False AND brands_overlap() to hold (when both sides have
      brand data) before a barcode link is learned.
  14. BARCODE-MATCH CONFLICT FLAGGING — even a *pre-existing* /
      dictionary barcode link is no longer trusted blindly. If a
      direct or bridged barcode match's weight or brand disagrees
      with the customer row, the result is now surfaced as
      `Matched (Barcode) – REVIEW: possible bad link` with confidence
      dropped to 60, instead of silently reporting 100%. This makes
      already-poisoned barcode links (from before this patch) visible
      in the output instead of invisible.

INSTALL:  pip install openai faiss-cpu pandas rapidfuzz openpyxl numpy httpx
"""

# ===========================================================================
# CONFIG
# ===========================================================================

import os as _os
from dotenv import load_dotenv
load_dotenv()

PROJECT_ROOT = r"C:\Users\HP\Desktop\zero1"
DATA_DIR     = PROJECT_ROOT + r"\data"
CACHE_DIR    = PROJECT_ROOT + r"\cache"
OUTPUT_DIR   = PROJECT_ROOT + r"\output"

INPUT_FILE   = r"C:\Users\HP\Downloads\categoriesandfinal_brand_standardized (1).xlsx"
INPUT_SHEET  = 0
OUTPUT_FILE  = OUTPUT_DIR + r"\ta_v99.xlsx"

# ★ Set this True ONCE after upgrading to this patch, to purge any
#   matches/backfilled barcode links that were poisoned by the old
#   (unvalidated) backfill logic. Then set back to False.
FORCE_RECOMPUTE = False

# ── ★ SAMPLE SIZE (for quick test runs) ─────────────────────────────────────
# None / 0  -> process ALL rows for every customer (normal full run).
# An int N  -> after dedup, cap each customer's row set to N rows before any
#              barcode/embedding/LLM work happens. Handy for a fast smoke
#              test of prompt/logic changes without burning the full budget.
SAMPLE_SIZE        = 500   # e.g. 200 for a quick test, None for full run
SAMPLE_RANDOM      = True    # True = random sample (reproducible via seed), False = first N rows
SAMPLE_RANDOM_SEED = 42

SOURCE_CUSTOMER_LABEL = "LULU"
NAME_COL           = "material_name"
BRAND_COL          = "brand_standardized"
BRAND_FALLBACK_COL = "brand_original"
CATEGORY_COL       = "lulu_category"
BARCODE_COL        = "barcode"
MATERIAL_CODE_COL  = "material_code"

BACKFILL_MIN_CONFIDENCE = 89
MIN_CONFIDENCE          = 75
WEIGHT_TOL              = 0.02
BRAND_FUZZY_THRESH      = 75
DEDUPE_CUSTOMER_BY_DESC = True
REQUIRE_BRAND_MATCH       = True
ALLOW_MATCH_WITHOUT_BRAND = True
OPENAI_API_KEY = _os.getenv("OPENAI_API_KEY")
EMBED_MODEL      = "text-embedding-3-small"
EMBED_BATCH_SIZE = 512
TOP_K_NEIGHBORS  = 5

AUTO_ACCEPT_THRESHOLD = 80
MIN_LLM_CONFIDENCE    = 60

LLM_MODEL               = "gpt-4o-mini"
LLM_TEMPERATURE          = 0
PASS_BARCODE_HINT_TO_LLM = True

COST_PRINT_INTERVAL_SEC = 30
MAX_BUDGET_USD           = None

# ── ★ RETRY / BACKOFF CONFIG ────────────────────────────────────────────────
API_REQUEST_TIMEOUT_SEC = 60      # per-request timeout (connect + read)
MAX_API_ATTEMPTS        = 5       # total attempts before giving up
BACKOFF_BASE_SEC        = 2.0     # delay = BACKOFF_BASE_SEC * 2**attempt
BACKOFF_MAX_SEC         = 60.0    # cap on any single sleep
BACKOFF_JITTER_FRAC     = 0.25    # +/- 25% random jitter on each sleep

# ── ★ CONCURRENCY CONFIG ─────────────────────────────────────────────────────
LLM_MAX_WORKERS = 4       # concurrent threads doing LLM adjudication calls
HTTP_POOL_SIZE  = 12      # underlying HTTP connection pool size for the OpenAI client

PRICE_PER_1K = {
    "text-embedding-3-small": {"input": 0.00002},
    "text-embedding-3-large": {"input": 0.00013},
    "gpt-4o-mini":            {"input": 0.00015, "output": 0.00060},
    "gpt-4o":                 {"input": 0.0025,  "output": 0.010},
}

# ── ★ ROW-LEVEL CACHE FILE ──────────────────────────────────────────────────
ROW_CACHE_FILE = _os.path.join(CACHE_DIR, "row_match_cache.json")

# ===========================================================================

import re, sys, time, warnings, hashlib, json, threading, random
import concurrent.futures
import httpx
import numpy as np
import pandas as pd
from collections import defaultdict
from pathlib import Path
from rapidfuzz import fuzz

warnings.filterwarnings("ignore")

for _d in (PROJECT_ROOT, DATA_DIR, CACHE_DIR, OUTPUT_DIR):
    _os.makedirs(_d, exist_ok=True)


# ===========================================================================
# ★ RETRY / EXPONENTIAL BACKOFF HELPER
# ===========================================================================

class APICallFailed(Exception):
    """Raised when call_with_backoff exhausts all attempts."""
    pass


def call_with_backoff(fn, *, what: str, max_attempts=MAX_API_ATTEMPTS,
                       base_delay=BACKOFF_BASE_SEC, max_delay=BACKOFF_MAX_SEC,
                       jitter_frac=BACKOFF_JITTER_FRAC):
    """
    Calls fn() and retries on ANY exception with exponential backoff + jitter.
    Logs every attempt (success or failure) so a stuck/failing call is
    visible in stdout instead of silently hanging the whole pipeline.

    Thread-safe: each call is independent, printing may interleave across
    threads but that's fine — it's log output, not shared state.

    fn        : zero-arg callable that performs the actual API call
    what      : short label for logging, e.g. "embed batch 4/35"
    Returns   : fn()'s return value on success.
    Raises    : APICallFailed if all attempts are exhausted.
    """
    last_exc = None
    for attempt in range(1, max_attempts + 1):
        t0 = time.time()
        try:
            result = fn()
            if attempt > 1:
                print(f"    ✓ {what} succeeded on attempt {attempt}/{max_attempts} "
                      f"({time.time()-t0:.1f}s)", flush=True)
            return result
        except Exception as e:
            last_exc = e
            elapsed = time.time() - t0
            print(f"    ⚠ {what} FAILED attempt {attempt}/{max_attempts} "
                  f"after {elapsed:.1f}s: {type(e).__name__}: {e}", flush=True)
            if attempt == max_attempts:
                break
            delay = min(max_delay, base_delay * (2 ** (attempt - 1)))
            jitter = delay * jitter_frac
            sleep_for = max(0.1, delay + random.uniform(-jitter, jitter))
            print(f"    … retrying {what} in {sleep_for:.1f}s "
                  f"(exponential backoff, attempt {attempt+1}/{max_attempts})", flush=True)
            time.sleep(sleep_for)

    raise APICallFailed(f"{what}: exhausted {max_attempts} attempts — last error: {last_exc}") from last_exc


# ===========================================================================
# ★ ROW-LEVEL PERSISTENT CACHE  (thread-safe)
# ===========================================================================

class RowMatchCache:
    """Persistent JSON dict:  row_hash -> match_result_dict.
    Loaded once at start, flushed to disk periodically and at end.

    ★ THREAD-SAFETY NOTE: get/put/flush/__len__ are all guarded by an
    RLock (re-entrant, since put() can internally call flush()).
    """

    def __init__(self, path=ROW_CACHE_FILE, flush_every=500):
        self.path = path
        self.flush_every = flush_every
        self._dirty = 0
        self.data = {}
        self._lock = threading.RLock()  # ★ re-entrant: put() may call flush()
        if not FORCE_RECOMPUTE and _os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
                print(f"  ✓ Row cache loaded: {len(self.data):,} entries → {path}")
            except Exception as e:
                print(f"  ⚠ Row cache corrupt ({e}) — starting fresh")
                self.data = {}

    @staticmethod
    def row_key(desc: str, brand, barcode) -> str:
        raw = f"{str(desc).strip()}|{str(brand).strip()}|{str(barcode).strip()}"
        return hashlib.sha256(raw.encode('utf-8', errors='ignore')).hexdigest()[:24]

    def get(self, key: str):
        with self._lock:
            return self.data.get(key)

    def put(self, key: str, result: dict):
        with self._lock:
            self.data[key] = result
            self._dirty += 1
            if self._dirty >= self.flush_every:
                self.flush()

    def flush(self):
        with self._lock:
            if self._dirty == 0:
                return
            try:
                with open(self.path, 'w', encoding='utf-8') as f:
                    json.dump(self.data, f, default=_json_safe_val)
                self._dirty = 0
            except Exception as e:
                print(f"  ⚠ Row cache flush failed: {e}")

    def __len__(self):
        with self._lock:
            return len(self.data)


def _json_safe_val(obj):
    if isinstance(obj, (np.integer,)):   return int(obj)
    if isinstance(obj, (np.floating,)):  return None if np.isnan(obj) else float(obj)
    if isinstance(obj, np.bool_):        return bool(obj)
    if isinstance(obj, float) and np.isnan(obj): return None
    try:
        if pd.isna(obj): return None
    except (TypeError, ValueError):
        pass
    return str(obj)


def _hash_list(items) -> str:
    h = hashlib.sha256()
    for it in items:
        h.update(str(it).encode('utf-8', errors='ignore'))
        h.update(b'\x00')
    return h.hexdigest()[:20]


def _hash_file(path) -> str:
    st = _os.stat(path)
    key = f"{path}|{st.st_size}|{int(st.st_mtime)}"
    return hashlib.sha256(key.encode('utf-8')).hexdigest()[:20]


def _json_safe(obj):
    if isinstance(obj, dict):
        return {k: _json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [_json_safe(v) for v in obj]
    if isinstance(obj, (np.integer,)):   return int(obj)
    if isinstance(obj, (np.floating,)):  return None if np.isnan(obj) else float(obj)
    if isinstance(obj, np.bool_):        return bool(obj)
    if isinstance(obj, float) and np.isnan(obj): return None
    try:
        if pd.isna(obj): return None
    except (TypeError, ValueError):
        pass
    return obj


# ===========================================================================
# COST TRACKER
# ===========================================================================

class CostTracker:
    def __init__(self, print_interval=COST_PRINT_INTERVAL_SEC, budget=MAX_BUDGET_USD):
        self.lock = threading.Lock()
        self.embed_calls = self.embed_tokens = 0
        self.llm_calls = self.llm_input_tokens = self.llm_output_tokens = 0
        self.rows_auto_accepted = self.rows_llm_escalated = 0
        self.rows_cache_hit = 0
        self.rows_hard_rejected = 0  # ★ NEW — LLM picked a candidate that failed brand/weight veto
        self.budget = budget
        self._start = time.time()
        self._stop_evt = threading.Event()
        self._thread = threading.Thread(target=self._loop, args=(print_interval,), daemon=True)
        self._thread.start()

    def add_embedding(self, n_tokens: int):
        with self.lock: self.embed_calls += 1; self.embed_tokens += n_tokens
    def add_llm(self, in_tok: int, out_tok: int):
        with self.lock: self.llm_calls += 1; self.llm_input_tokens += in_tok; self.llm_output_tokens += out_tok
    def note_auto_accept(self):
        with self.lock: self.rows_auto_accepted += 1
    def note_llm_escalation(self):
        with self.lock: self.rows_llm_escalated += 1
    def note_cache_hit(self):
        with self.lock: self.rows_cache_hit += 1
    def note_hard_reject(self):
        with self.lock: self.rows_hard_rejected += 1

    def cost_usd(self) -> float:
        with self.lock:
            ep = PRICE_PER_1K.get(EMBED_MODEL, {}).get("input", 0.0)
            li = PRICE_PER_1K.get(LLM_MODEL, {}).get("input", 0.0)
            lo = PRICE_PER_1K.get(LLM_MODEL, {}).get("output", 0.0)
            return (self.embed_tokens/1000)*ep + (self.llm_input_tokens/1000)*li + (self.llm_output_tokens/1000)*lo

    def budget_exceeded(self) -> bool:
        return self.budget is not None and self.cost_usd() >= self.budget

    def print_summary(self):
        el = time.time() - self._start
        with self.lock:
            ec, et = self.embed_calls, self.embed_tokens
            lc, lit, lot = self.llm_calls, self.llm_input_tokens, self.llm_output_tokens
            auto, esc, ch, hr = self.rows_auto_accepted, self.rows_llm_escalated, self.rows_cache_hit, self.rows_hard_rejected
        cost = self.cost_usd()
        bud = f" / ${self.budget:.2f}" if self.budget else ""
        print(f"[COST @ {el:6.0f}s] emb:{ec:,}c/{et:,}t | LLM:{lc:,}c/{lit:,}+{lot:,}t | "
              f"auto:{auto:,} llm:{esc:,} cache:{ch:,} hard_reject:{hr:,} | ${cost:.4f}{bud}", flush=True)

    def _loop(self, interval):
        while not self._stop_evt.wait(interval): self.print_summary()
    def stop(self):
        self._stop_evt.set(); self.print_summary()


# ===========================================================================
# PARSING / BRAND / WEIGHT HELPERS  (unchanged domain logic)
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
    r'\bGLD\b': 'GOLD', r'\bGLDN\b': 'GOLDEN',
    r'\bLNG\b': 'LONG', r'\bSML\b': 'SMALL', r'\bMED\b': 'MEDIUM',
    r'\bLRG\b': 'LARGE', r'\bXL\b': 'EXTRALARGE',
    r'\bORIG\b': 'ORIGINAL', r'\bORGNL\b': 'ORIGINAL',
    r'\bTRD\b': 'TRADITIONAL', r'\bSTD\b': 'STANDARD',
    r'\bPRM\b': 'PREMIUM', r'\bPRMM\b': 'PREMIUM',
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
    'AL', 'EL', 'BIN', 'IBN', 'ABU', 'UM', 'AA',
}

HIER_STOPWORDS = {
    'AND', 'OR', 'THE', 'OF', 'IN', 'FOR', 'WITH', 'A', 'AN',
    'FRESH', 'FROZEN', 'CHILLED', 'CANNED', 'DRIED',
    'WHOLE', 'HALF', 'SLICED', 'MIXED', 'ASSORTED',
    'PRODUCT', 'PRODUCTS', 'ITEM', 'OTHER', 'MISC',
}
COLOR_TYPE_GROUPS = [
    {'WHITE', 'BROWN'}, {'FULL', 'SKIMMED', 'SEMI'},
    {'SALTED', 'UNSALTED'}, {'SWEETENED', 'UNSWEETENED'},
    {'REGULAR', 'DIET', 'ZERO'}, {'LARGE', 'MEDIUM', 'SMALL'},
    {'CANOLA', 'CORN', 'SUNFLOWER', 'OLIVE', 'VEGETABLE', 'SESAME', 'SOYBEAN'},
    {'CHOCOLATE', 'BUTTER', 'VANILLA', 'COCONUT', 'OATMEAL', 'GINGER', 'LEMON', 'ALMOND'},
    {'JASMINE', 'BASMATI', 'SELLA', 'PARBOILED'},
    {'FULLFAT', 'LOWFAT', 'NONFAT', 'SKIMMEDMILK'},
]

PACK_RE = re.compile(
    r'\b(\d+\s*[Xx]\s*\d+(?:\s*[Xx]\s*\d+(?:\.\d+)?)?'
    r'(?:\s*(?:KG|KGS|G|GR|GM|GMS|ML|L|LT|LTR|LTRS|MG|S))?)\b', re.I)
SIZE_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*(KG|KGS|G|GR|GM|GMS|GRM|GRMS|GRAMS?|MG|'
    r'L|LT|LTR|LTRS|LITRE|LITER|ML|CL|CC|S)\b', re.I)
GLUED_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*(KG|KGS|G|GR|GM|GMS|ML|L|LT|LTR|MG|S)\s*X\s*(\d+)\b', re.I)
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
CODE_RE = re.compile(r'\(\s*\d{5,}\s*\)', re.I)


def strip_product_codes(text: str) -> str:
    return CODE_RE.sub(' ', text)


def color_type_penalty(name_a: str, name_b: str) -> bool:
    tokens_a = set(TOKEN_RE.findall(name_a.upper()))
    tokens_b = set(TOKEN_RE.findall(name_b.upper()))
    for group in COLOR_TYPE_GROUPS:
        hits_a = tokens_a & group
        hits_b = tokens_b & group
        if hits_a and hits_b and hits_a != hits_b:
            return True
    return False


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
        if unit not in UNIT_TABLE: continue
        mult, kind = UNIT_TABLE[unit]
        try: sizes_found.append((float(m.group(1)) * mult, kind))
        except ValueError: pass
    for m in GLUED_RE.finditer(s):
        unit = m.group(2).upper()
        if unit not in UNIT_TABLE: continue
        mult, kind = UNIT_TABLE[unit]
        try: sizes_found.append((float(m.group(1)) * mult, kind))
        except ValueError: pass
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
        if gm: pack_str = gm.group(3)
    name = s
    name = GLUED_RE.sub(' ', name)
    name = PACK_RE.sub(' ', name)
    name = SIZE_RE.sub(' ', name)
    name = re.sub(r'\bX\s*\d+\b', ' ', name)
    name = re.sub(r'\bWS\b', ' ', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return (name, weight_base, weight_kind, pack_str)


def normalize_barcode(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, float): val = f'{val:.0f}'
    s = str(val).strip()
    if s == '' or s.lower() in ('nan', 'none', '0', '0.0'): return None
    if 'e' in s.lower():
        try: s = f'{float(s):.0f}'
        except ValueError: pass
    if s.endswith('.0'): s = s[:-2]
    s = re.sub(r'\D', '', s)
    if not s or set(s) == {'0'} or len(s) < 6: return None
    return s


def normalize_name_key(val):
    if val is None or (isinstance(val, float) and pd.isna(val)): return None
    s = str(val).strip().upper()
    s = re.sub(r'\s+', ' ', s)
    return s if s else None


def normalise_brand_str(raw) -> frozenset:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)): return frozenset()
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
        if t in NON_BRAND or t.isdigit() or len(t) < 2: continue
        out.add(BRAND_ABBREV.get(t, t))
    return frozenset(out)


def brand_tokens_from_desc(clean_name: str) -> frozenset:
    if not clean_name: return frozenset()
    for w in clean_name.split()[:2]:
        if w in NON_BRAND or w.isdigit() or len(w) < 3: continue
        return frozenset({BRAND_ABBREV.get(w, w)})
    return frozenset()


def brands_overlap(a: frozenset, b: frozenset) -> bool:
    if not a or not b: return False
    for x in a:
        for y in b:
            if x == y: return True
            if len(x) >= 5 and len(y) >= 5:
                if x in y or y in x: return True
                if x[:5] == y[:5]: return True
    return False


def brand_label(src_brand_raw, cust_brand_raw, threshold, alias_dict):
    sb = str(src_brand_raw).strip().lower() if src_brand_raw else ''
    cb = str(cust_brand_raw).strip().lower() if cust_brand_raw else ''
    if not sb or not cb or sb in ('nan', 'none') or cb in ('nan', 'none'): return 'N/A'
    if sb == cb: return 'Exact'
    score = fuzz.token_sort_ratio(sb, cb)
    if score >= threshold:
        if sb not in alias_dict:
            alias_dict[sb] = {'Customer Brand': cb, 'Fuzzy Score': score}
        return 'Fuzzy'
    return 'Mismatch'


def weights_match(w1, w2, tol=WEIGHT_TOL):
    if w1 is None or w2 is None: return None
    if (isinstance(w1, float) and np.isnan(w1)) or (isinstance(w2, float) and np.isnan(w2)): return None
    if w1 <= 0 or w2 <= 0: return None
    return (min(w1, w2) / max(w1, w2)) >= (1.0 - tol)


def pack_match_bonus(p1, p2):
    return 3.0 if (p1 and p2 and p1 == p2) else 0.0


def weight_match_label(src_w, cust_w):
    if src_w is None and cust_w is None: return 'N/A – no weight on either side'
    if src_w is not None and cust_w is not None:
        return (f'Matched – both {src_w:.1f}' if weights_match(src_w, cust_w)
                else f'Mismatch – source {src_w:.1f} vs customer {cust_w:.1f}')
    return f'Source only – {src_w:.1f}' if src_w is not None else f'Customer only – {cust_w:.1f}'


def category_bonus(cust_tokens, src_h1, pts=4.0, cap=10.0):
    if not cust_tokens: return 0.0
    hier_text = ''
    if src_h1 and not (isinstance(src_h1, float) and pd.isna(src_h1)):
        hier_text += str(src_h1).upper()
    if not hier_text.strip(): return 0.0
    hier_tokens = {t for t in TOKEN_RE.findall(hier_text) if len(t) >= 3 and t not in HIER_STOPWORDS}
    return min(cap, len(cust_tokens & hier_tokens) * pts)


# ===========================================================================
# PARQUET INPUT CACHE
# ===========================================================================

def load_input_dataframe() -> pd.DataFrame:
    src_path = Path(INPUT_FILE)
    if not src_path.exists():
        print(f"ERROR: input file not found: {INPUT_FILE}"); sys.exit(1)
    file_key     = _hash_file(str(src_path))
    parquet_path = _os.path.join(DATA_DIR, f"input_{file_key}.parquet")
    if not FORCE_RECOMPUTE and _os.path.exists(parquet_path):
        try:
            print(f"  ✓ Parquet cache hit → {parquet_path}")
            return pd.read_parquet(parquet_path)
        except Exception as e:
            print(f"  ⚠ Parquet unreadable ({e}) — reconverting")
    print("  Converting Excel → Parquet …")
    if INPUT_SHEET is not None:
        df = pd.read_excel(str(src_path), sheet_name=INPUT_SHEET, dtype={BARCODE_COL: str})
    elif str(src_path).lower().endswith('.xlsx'):
        df = pd.read_excel(str(src_path), dtype={BARCODE_COL: str})
    else:
        df = pd.read_csv(str(src_path), low_memory=False, dtype={BARCODE_COL: str})
    try:
        df.to_parquet(parquet_path, index=False)
        print(f"  ✓ Cached → {parquet_path}")
    except Exception as e:
        print(f"  ⚠ Parquet write failed ({e})")
    return df


# ===========================================================================
# ★ SHARED HTTPX CLIENT (bigger connection pool for concurrent threads)
# ===========================================================================

def _make_http_client() -> httpx.Client:
    return httpx.Client(
        limits=httpx.Limits(
            max_connections=HTTP_POOL_SIZE,
            max_keepalive_connections=HTTP_POOL_SIZE,
        )
    )


# ===========================================================================
# OPENAI EMBEDDINGS + FAISS
# ===========================================================================

class OpenAIFaissIndex:
    def __init__(self, source_names: list, cost_tracker):
        import faiss
        from openai import OpenAI
        self.faiss = faiss
        self.client = OpenAI(
            api_key=OPENAI_API_KEY,
            timeout=API_REQUEST_TIMEOUT_SEC,
            max_retries=0,
            http_client=_make_http_client(),
        )
        self.cost_tracker = cost_tracker

        cache_key = _hash_list(source_names) + '_' + EMBED_MODEL.replace('/', '_')
        npy_path  = _os.path.join(CACHE_DIR, f"src_emb_{cache_key}.npy")
        json_path = _os.path.join(CACHE_DIR, f"src_emb_{cache_key}.json")

        self.src_emb = None
        if not FORCE_RECOMPUTE and _os.path.exists(npy_path):
            try:
                emb = np.load(npy_path)
                if emb.shape[0] == len(source_names):
                    self.src_emb = emb
                    print(f"  ✓ Source embedding cache hit → {npy_path}")
            except Exception as e:
                print(f"  ⚠ .npy corrupt ({e})")

        if self.src_emb is None and not FORCE_RECOMPUTE and _os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f: raw = json.load(f)
                emb = np.array(raw, dtype=np.float32)
                if emb.shape[0] == len(source_names):
                    self.src_emb = emb
                    print(f"  ✓ Source embedding cache hit (.json) → {json_path}")
                    try: np.save(npy_path, self.src_emb)
                    except: pass
            except Exception as e:
                print(f"  ⚠ .json corrupt ({e})")

        if self.src_emb is None:
            print(f"  Embedding {len(source_names):,} source names via {EMBED_MODEL} …")
            self.src_emb = self._embed_texts(source_names, label="source")
            try:
                np.save(npy_path, self.src_emb)
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(self.src_emb.tolist(), f)
                print(f"  ✓ Cached → {npy_path}")
            except Exception as e:
                print(f"  ⚠ Cache write failed: {e}")

        dim = self.src_emb.shape[1]
        self.index = faiss.IndexFlatIP(dim)
        self.index.add(np.ascontiguousarray(self.src_emb.astype('float32')))
        print(f"  ✓ FAISS index: {self.index.ntotal:,} vectors, dim={dim}")

    def _embed_texts(self, texts: list, label: str = "") -> np.ndarray:
        all_vecs = []
        total_batches = (len(texts) + EMBED_BATCH_SIZE - 1) // EMBED_BATCH_SIZE

        for start in range(0, len(texts), EMBED_BATCH_SIZE):
            batch = texts[start:start + EMBED_BATCH_SIZE]
            batch = [t if t.strip() else " " for t in batch]
            batch_num = start // EMBED_BATCH_SIZE + 1

            print(f"    [{label}] batch {batch_num}/{total_batches} "
                  f"({len(batch)} items) — requesting…", flush=True)

            def _do_call(batch=batch):
                return self.client.embeddings.create(model=EMBED_MODEL, input=batch)

            resp = call_with_backoff(
                _do_call,
                what=f"embed batch {batch_num}/{total_batches} [{label}]",
            )

            vecs = np.array([d.embedding for d in resp.data], dtype=np.float32)
            all_vecs.append(vecs)
            if resp.usage:
                self.cost_tracker.add_embedding(resp.usage.total_tokens)

            print(f"    ✓ [{label}] batch {batch_num}/{total_batches} done "
                  f"({min(start+EMBED_BATCH_SIZE, len(texts)):,}/{len(texts):,} total)", flush=True)

        emb = np.vstack(all_vecs)
        norms = np.linalg.norm(emb, axis=1, keepdims=True)
        norms[norms == 0] = 1.0
        return emb / norms

    def embed_queries_cached(self, texts: list, cache_label: str) -> np.ndarray:
        cache_key = _hash_list(texts) + '_' + EMBED_MODEL.replace('/', '_')
        safe = re.sub(r'[^A-Za-z0-9_-]', '_', cache_label)[:40]
        npy_path = _os.path.join(CACHE_DIR, f"qemb_{safe}_{cache_key}.npy")
        if not FORCE_RECOMPUTE and _os.path.exists(npy_path):
            try:
                emb = np.load(npy_path)
                if emb.shape[0] == len(texts):
                    print(f"  ✓ Query embedding cache hit → {npy_path}")
                    return emb
            except: pass
        emb = self._embed_texts(texts, label=safe)
        try: np.save(npy_path, emb)
        except: pass
        return emb

    def search(self, query_emb: np.ndarray, k=TOP_K_NEIGHBORS):
        q = np.ascontiguousarray(query_emb.astype('float32'))
        return self.index.search(q, k)


def cosine_to_pct(sim):
    return max(0.0, min(100.0, (sim + 1.0) / 2.0 * 100.0))


# ===========================================================================
# RERANK
# ===========================================================================

def rerank_candidates(
    raw_scores, cand_idxs,
    c_weight, c_kind, c_pack, c_brands, cust_tokens, cust_name_str,
    src_weight, src_kind, src_pack, src_brands, src_h1, src_name,
):
    out = []
    for sim, si in zip(raw_scores, cand_idxs):
        if si < 0: continue
        score = cosine_to_pct(sim)
        score = min(100.0, score + category_bonus(cust_tokens, src_h1[si]))
        if color_type_penalty(cust_name_str, src_name[si]):
            out.append((si, 0.0, False)); continue
        wm = weights_match(c_weight, src_weight[si])
        if wm is False: score = 0.0
        elif wm is True:
            score = min(100.0, score + 5.0)
            if c_kind and src_kind[si] and c_kind != src_kind[si]:
                score = max(0.0, score - 20.0)
        score = min(100.0, score + pack_match_bonus(c_pack, src_pack[si]))
        brand_ok = brands_overlap(c_brands, src_brands[si]) or (
            (not REQUIRE_BRAND_MATCH) or (ALLOW_MATCH_WITHOUT_BRAND and not c_brands and not src_brands[si]))
        out.append((si, score, brand_ok))
    out.sort(key=lambda t: t[1], reverse=True)
    return out


# ===========================================================================
# ★★ LLM ADJUDICATION — brand + weight are now explicit HARD RULES
# ===========================================================================

_SYSTEM_PROMPT = """\
Product-match adjudicator for Qatar FMCG retail. Source catalogue: LULU Hypermarket.
Incoming data is messy: abbreviated units (PKT, FZ), inconsistent brand spelling, Arabic \
transliterations, glued sizes (6X330ML). This is normal surface noise — don't penalize it.

Given 1 customer row + up to 5 LULU candidates, decide if any candidate is the SAME product.

HARD RULES — a candidate is NOT a match (match_index must skip it) if:
  1. WEIGHT/PACK differs. Same item at a different weight/pack size is a DIFFERENT product.
  2. BRAND differs. Brand identity always outranks category or description similarity.
     Example: "Seara" chicken is NOT "Asaffa" chicken. "Al Waha" is NOT "Al Naeem".
     A brand match IS allowed across spelling/transliteration variants of the SAME brand
     (e.g. "Ahmad Tea" / "Ahmed Tea"), or a private-label / generic customer row with no
     stated brand — but never across two different named brands.

If NO candidate satisfies BOTH rules, return match_index: null, even if the category or
general description looks close. A close-looking wrong-brand or wrong-weight candidate is
a worse outcome than returning no match.

Barcode hint (if present) is supporting evidence only (it didn't match directly on its own).

Reply ONLY with JSON: {"match_index": <1-based or null>, "confidence": <0-100>, "reason": "<short>"}"""


def _build_llm_user_prompt(cust_desc, cust_brand, cust_weight, cust_pack, barcode_hint, candidates):
    """Compact single-line format — saves tokens per call."""
    parts = [f"CUST: {cust_desc} | brand:{cust_brand or '-'} | wt:{cust_weight or '?'} | pack:{cust_pack or '-'}"]
    if PASS_BARCODE_HINT_TO_LLM and barcode_hint:
        parts[0] += f" | bc_hint:{barcode_hint}"
    parts.append("CANDS:")
    for i, c in enumerate(candidates, 1):
        parts.append(f" {i}. {c['name']} | {c['brand'] or '-'} | wt:{c['weight'] or '?'} | "
                     f"cat:{c['category'] or '-'} | bc:{c['barcode'] or '-'} | s:{c['score']:.0f}")
    return "\n".join(parts)


def llm_adjudicate(client, cust_desc, cust_brand, cust_weight, cust_pack, barcode_hint, candidates, cost_tracker,
                    row_label: str = ""):
    """
    Thread-safe by design: takes no shared mutable state except
    `cost_tracker` (lock-protected) and the OpenAI client (thread-safe
    HTTP client). Multiple worker threads can call this concurrently.

    NOTE: this function returns exactly what the LLM said. It does NOT
    itself enforce brand/weight — that veto happens in the caller
    (match_customer, PHASE 2) against the pre-computed candidate
    validity flags, so the veto logic lives in one deterministic place
    instead of depending on the model always following the prompt.
    """
    user_prompt = _build_llm_user_prompt(cust_desc, cust_brand, cust_weight, cust_pack, barcode_hint, candidates)

    def _do_call():
        return client.chat.completions.create(
            model=LLM_MODEL, temperature=LLM_TEMPERATURE,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": _SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
        )

    try:
        resp = call_with_backoff(_do_call, what=f"LLM adjudicate [{row_label}]")
    except APICallFailed as e:
        print(f"  ⚠ LLM call permanently failed: {e}", flush=True)
        return None, 0, f"llm_error: {e}"

    if resp.usage:
        cost_tracker.add_llm(resp.usage.prompt_tokens, resp.usage.completion_tokens)
    raw = resp.choices[0].message.content
    try:
        parsed = json.loads(raw)
        idx = parsed.get("match_index")
        conf = int(parsed.get("confidence", 0))
        reason = str(parsed.get("reason", ""))
        if idx is not None:
            idx = int(idx)
            if not (1 <= idx <= len(candidates)): idx = None
        return idx, conf, reason
    except Exception as e:
        print(f"  ⚠ LLM JSON parse error ({e}): {raw[:200]!r}")
        return None, 0, "unparseable_llm_response"


# ===========================================================================
# CUSTOMER FRAME BUILDER
# ===========================================================================

def get_customer_frame(df_all, label, barcode_to_src=None):
    sub = df_all[df_all['source_customer'] == label].copy()
    sub = sub.dropna(subset=[NAME_COL])

    sub['_dedupe_key'] = sub[NAME_COL].astype(str).map(normalize_name_key)

    dup_counts = {}
    if DEDUPE_CUSTOMER_BY_DESC:
        before = len(sub)
        dup_counts = sub.groupby('_dedupe_key').size().to_dict()

        if barcode_to_src and BARCODE_COL and BARCODE_COL in sub.columns:
            def _pri(bc_raw):
                bc = normalize_barcode(bc_raw)
                return 0 if (bc and bc in barcode_to_src) else 1
            sub['_dp'] = sub[BARCODE_COL].apply(_pri)
            sub = sub.sort_values('_dp', kind='stable').drop_duplicates(subset=['_dedupe_key'], keep='first').drop(columns=['_dp'])
        else:
            sub = sub.drop_duplicates(subset=['_dedupe_key'])
        removed = before - len(sub)
        if removed: print(f"  [{label}] Dropped {removed:,} duplicate row(s) (normalized key match) — each unique item matches only once")

    # ── ★ SAMPLE_SIZE — cap rows per customer for quick test runs ──
    # Applied AFTER dedup so barcode/embedding/LLM work is capped too.
    if SAMPLE_SIZE and len(sub) > SAMPLE_SIZE:
        before_n = len(sub)
        if SAMPLE_RANDOM:
            sub = sub.sample(n=SAMPLE_SIZE, random_state=SAMPLE_RANDOM_SEED)
        else:
            sub = sub.head(SAMPLE_SIZE)
        print(f"  [{label}] ★ SAMPLE_SIZE active — using {len(sub):,}/{before_n:,} rows "
              f"({'random' if SAMPLE_RANDOM else 'first-N'}, seed={SAMPLE_RANDOM_SEED if SAMPLE_RANDOM else '-'})")

    sub = sub.reset_index(drop=True)
    out = pd.DataFrame({'desc': sub[NAME_COL].astype(str), 'brand': sub['_brand_resolved']})
    out['barcode'] = sub[BARCODE_COL].values if (BARCODE_COL and BARCODE_COL in sub.columns) else None
    out['material_code'] = sub[MATERIAL_CODE_COL].values if (MATERIAL_CODE_COL and MATERIAL_CODE_COL in sub.columns) else None
    out['dup_count'] = sub['_dedupe_key'].map(dup_counts).fillna(1).astype(int) if dup_counts else 1
    return out


# ===========================================================================
# ★★ MATCH ONE CUSTOMER GROUP
#    (row-level cache + query dedup + threaded LLM + brand/weight veto
#     + validated barcode backfill + barcode-match conflict flagging)
# ===========================================================================

def match_customer(
    label, cust,
    src_name, src_weight, src_kind, src_pack, src_brands, src_h1,
    src_descraw, src_brand_display, src_barcode_display, src_material_code_display,
    barcode_to_src, name_key_to_src,
    faiss_index, openai_client, cost_tracker, row_cache,
):
    t0 = time.time()
    print(f"\n{'='*60}\n  Customer: {label}\n{'='*60}")
    print(f"  Rows: {len(cust):,}")

    def _maybe_backfill(cust_barcode_raw, src_i, method, conf_int,
                         cust_weight_val=None, cust_brands_val=None):
        """
        ★★ VALIDATED backfill. A barcode link is only allowed to be
        learned (i.e. trusted as ground truth for every future
        customer row sharing that barcode) if:
          - it's a direct or bridged barcode hit (already-verified
            identity, method in ('Barcode','Barcode (Bridged)')), OR
          - it's a high-confidence Embedding/LLM match AND that match
            does not conflict on weight (weights_match() != False)
            AND does not conflict on brand (brands_overlap() holds,
            when both sides actually have brand tokens).
        This stops a single bad high-confidence match from silently
        poisoning barcode_to_src and being inherited as a "100%
        Matched (Barcode)" result by every later customer row with
        the same barcode.
        """
        cust_bc_norm = normalize_barcode(cust_barcode_raw)
        if not cust_bc_norm: return None, False

        is_high = method in ('Embedding', 'LLM') and conf_int is not None and conf_int >= BACKFILL_MIN_CONFIDENCE
        is_bridge = method in ('Barcode (Bridged)', 'Barcode')

        if is_high:
            wm = weights_match(cust_weight_val, src_weight[src_i])
            if wm is False:
                return None, False  # ★ weight conflict — refuse to learn this link
            sb = src_brands[src_i]
            if cust_brands_val and sb and not brands_overlap(cust_brands_val, sb):
                return None, False  # ★ brand conflict — refuse to learn this link

        if not (is_high or is_bridge): return None, False
        newly = cust_bc_norm not in barcode_to_src
        barcode_to_src[cust_bc_norm] = src_i
        if not src_barcode_display[src_i]: src_barcode_display[src_i] = cust_barcode_raw
        return cust_barcode_raw, newly

    cust['_barcode'] = cust['barcode'].apply(normalize_barcode)

    # ── step 1: barcode pass ──
    barcode_matched_cis = {}
    for ci, bc in enumerate(cust['_barcode'].tolist()):
        if bc and bc in barcode_to_src:
            barcode_matched_cis[ci] = (barcode_to_src[bc], 'Barcode')
    n_bridged = 0
    if name_key_to_src:
        for ci in range(len(cust)):
            if ci in barcode_matched_cis: continue
            key = normalize_name_key(cust.at[ci, 'desc'])
            if key and key in name_key_to_src:
                barcode_matched_cis[ci] = (name_key_to_src[key], 'Barcode (Bridged)')
                n_bridged += 1
    n_direct = len(barcode_matched_cis) - n_bridged
    print(f"  Barcode: {n_direct:,} direct + {n_bridged:,} bridged = {len(barcode_matched_cis):,}")

    # ── parse all customer rows ──
    parsed_cust = [parse_description(t) for t in cust['desc']]
    cust_name   = [p[0] for p in parsed_cust]
    cust_weight = [p[1] for p in parsed_cust]
    cust_kind   = [p[2] for p in parsed_cust]
    cust_pack   = [p[3] for p in parsed_cust]
    cust_brands = [normalise_brand_str(b) | brand_tokens_from_desc(nm)
                   for b, nm in zip(cust['brand'], cust_name)]
    cust_token_sets = [{t for t in TOKEN_RE.findall(nm) if len(t) >= 3 and t not in HIER_STOPWORDS}
                       for nm in cust_name]

    # ── step 2: check row cache for remaining rows ──
    remaining_cis = [ci for ci in range(len(cust)) if ci not in barcode_matched_cis]
    cache_hit_cis = {}
    still_need    = []

    for ci in remaining_cis:
        rk = RowMatchCache.row_key(cust.at[ci, 'desc'], cust.at[ci, 'brand'], cust.at[ci, 'barcode'])
        cached = row_cache.get(rk)
        if cached is not None:
            cache_hit_cis[ci] = cached
            cost_tracker.note_cache_hit()
        else:
            still_need.append(ci)

    print(f"  Row cache hits: {len(cache_hit_cis):,} | Still need embedding: {len(still_need):,}")

    # ── step 3: deduplicate query texts before embedding ──
    embedded_results = {}
    if still_need:
        unique_texts = {}
        for ci in still_need:
            txt = cust_name[ci] if cust_name[ci] else cust['desc'][ci]
            unique_texts.setdefault(txt, []).append(ci)

        dedup_texts = list(unique_texts.keys())
        n_saved = len(still_need) - len(dedup_texts)
        if n_saved > 0:
            print(f"  ★ Query dedup: {len(still_need):,} rows → {len(dedup_texts):,} unique texts (saved {n_saved:,} embeddings)")

        print(f"  Embedding {len(dedup_texts):,} unique queries via {EMBED_MODEL} …")
        query_emb = faiss_index.embed_queries_cached(dedup_texts, label)
        raw_scores_all, raw_idxs_all = faiss_index.search(query_emb, k=TOP_K_NEIGHBORS)

        dedup_faiss = {}
        for di in range(len(dedup_texts)):
            dedup_faiss[di] = (raw_scores_all[di], raw_idxs_all[di])

        n_auto = 0
        n_esc = 0
        llm_tasks = []

        # ── PHASE 1: rerank every row (fast, CPU-only, sequential) ──
        for di, txt in enumerate(dedup_texts):
            cis_for_text = unique_texts[txt]
            scores_row, idxs_row = dedup_faiss[di]

            for ci in cis_for_text:
                if cost_tracker.budget_exceeded():
                    print("  ⚠ BUDGET reached — stopping.")
                    break

                reranked = rerank_candidates(
                    scores_row, idxs_row,
                    cust_weight[ci], cust_kind[ci], cust_pack[ci], cust_brands[ci],
                    cust_token_sets[ci], cust_name[ci],
                    src_weight, src_kind, src_pack, src_brands, src_h1, src_name,
                )
                if not reranked:
                    result = {'method': 'none', 'src_i': None, 'score': 0, 'reason': 'no_candidates'}
                    embedded_results[ci] = result
                    rk = RowMatchCache.row_key(cust.at[ci, 'desc'], cust.at[ci, 'brand'], cust.at[ci, 'barcode'])
                    row_cache.put(rk, result)
                    continue

                best_si, best_score, best_brand_ok = reranked[0]

                if best_brand_ok and best_score >= AUTO_ACCEPT_THRESHOLD:
                    result = {'method': 'Embedding', 'src_i': int(best_si), 'score': float(best_score), 'reason': 'ok'}
                    embedded_results[ci] = result
                    cost_tracker.note_auto_accept()
                    n_auto += 1
                    rk = RowMatchCache.row_key(cust.at[ci, 'desc'], cust.at[ci, 'brand'], cust.at[ci, 'barcode'])
                    row_cache.put(rk, result)
                    continue

                # ★ queue for concurrent LLM adjudication (not called yet)
                cost_tracker.note_llm_escalation()
                n_esc += 1
                top_candidates = reranked[:TOP_K_NEIGHBORS]
                cand_payload = [{
                    'name': src_name[si], 'brand': src_brand_display[si],
                    'weight': src_weight[si], 'category': src_h1[si],
                    'barcode': src_barcode_display[si], 'score': sc,
                } for si, sc, _ok in top_candidates]

                # ★★ NEW: pre-computed brand_ok / weight_ok per candidate,
                #    same order as cand_payload / top_candidates — this is
                #    what the LLM's choice gets vetoed against in PHASE 2,
                #    independent of what the LLM claims about its own pick.
                cand_validity = []
                for si, sc, brand_ok in top_candidates:
                    wm = weights_match(cust_weight[ci], src_weight[si])
                    cand_validity.append({
                        'brand_ok': bool(brand_ok),
                        'weight_ok': (wm is not False),  # None (no data either side) is allowed through
                    })

                barcode_hint = cust.at[ci, 'barcode'] if PASS_BARCODE_HINT_TO_LLM else None
                row_key = RowMatchCache.row_key(cust.at[ci, 'desc'], cust.at[ci, 'brand'], cust.at[ci, 'barcode'])

                llm_tasks.append({
                    'ci': ci,
                    'cust_name': cust_name[ci] or cust.at[ci, 'desc'],
                    'cust_brand': cust['brand'][ci],
                    'cust_weight': cust_weight[ci],
                    'cust_pack': cust_pack[ci],
                    'barcode_hint': barcode_hint,
                    'top_candidates': top_candidates,
                    'cand_payload': cand_payload,
                    'cand_validity': cand_validity,   # ★ NEW
                    'row_key': row_key,
                })

        # ── PHASE 2: run all queued LLM adjudications concurrently ──
        if llm_tasks:
            print(f"  ★ Dispatching {len(llm_tasks):,} LLM adjudications "
                  f"across {LLM_MAX_WORKERS} workers …", flush=True)

            done_count = 0
            print_lock = threading.Lock()

            def _run_task(task):
                idx, conf, reason = llm_adjudicate(
                    openai_client, task['cust_name'], task['cust_brand'],
                    task['cust_weight'], task['cust_pack'], task['barcode_hint'],
                    task['cand_payload'], cost_tracker,
                    row_label=f"{label} row {task['ci']}",
                )
                return task, idx, conf, reason

            with concurrent.futures.ThreadPoolExecutor(max_workers=LLM_MAX_WORKERS) as executor:
                futures = [executor.submit(_run_task, task) for task in llm_tasks]

                for fut in concurrent.futures.as_completed(futures):
                    task, match_idx, llm_conf, reason = fut.result()
                    ci = task['ci']
                    top_candidates = task['top_candidates']
                    cand_validity  = task['cand_validity']

                    if match_idx is None or llm_conf < MIN_LLM_CONFIDENCE:
                        result = {'method': 'LLM', 'src_i': None, 'score': int(llm_conf),
                                  'reason': reason or 'llm_no_match'}
                    else:
                        v = cand_validity[match_idx - 1]
                        # ★★ POST-LLM HARD VETO — never trust the LLM's pick
                        # blindly; it must also pass the same deterministic
                        # brand/weight checks the embedding auto-accept path uses.
                        if not v['weight_ok']:
                            cost_tracker.note_hard_reject()
                            result = {'method': 'LLM', 'src_i': None, 'score': 0,
                                      'reason': f'hard_reject_weight_conflict (llm conf={llm_conf}: {reason})'}
                        elif not v['brand_ok']:
                            cost_tracker.note_hard_reject()
                            result = {'method': 'LLM', 'src_i': None, 'score': 0,
                                      'reason': f'hard_reject_brand_conflict (llm conf={llm_conf}: {reason})'}
                        else:
                            chosen_si = top_candidates[match_idx - 1][0]
                            result = {'method': 'LLM', 'src_i': int(chosen_si), 'score': int(llm_conf), 'reason': reason}

                    embedded_results[ci] = result
                    row_cache.put(task['row_key'], result)   # ★ thread-safe write

                    done_count += 1
                    if done_count % 200 == 0:
                        with print_lock:
                            print(f"  progress {n_auto + done_count}/{len(still_need)} "
                                  f"(auto:{n_auto} llm_done:{done_count}/{len(llm_tasks)}) "
                                  f"({time.time()-t0:.1f}s)", flush=True)

        print(f"  Embedding auto: {n_auto:,} | LLM: {n_esc:,}")

    # merge cache hits into embedded_results
    for ci, cached in cache_hit_cis.items():
        embedded_results[ci] = cached

    # ── assemble output ──
    matched_rows, unmatched_rows, alias_dict, matched_src_indices = [], [], {}, set()
    n_backfilled = 0

    for ci in range(len(cust)):
        cust_desc_raw   = cust.at[ci, 'desc']
        cust_brand_raw  = cust.at[ci, 'brand']
        cust_barcode_raw = cust.at[ci, 'barcode']
        cust_mc_raw     = cust.at[ci, 'material_code']
        cust_wt, cust_pk = cust_weight[ci], cust_pack[ci]

        dup_count = int(cust.at[ci, 'dup_count']) if 'dup_count' in cust.columns else 1
        base = {
            'Source File': label, 'Customer Description': cust_desc_raw,
            'Customer Brand': cust_brand_raw, 'Customer Material Code': cust_mc_raw,
            'Customer Barcode': cust_barcode_raw, 'Customer Weight g/ml': cust_wt, 'Customer Pack': cust_pk,
            'Duplicate Rows In Source': dup_count,
        }

        if ci in barcode_matched_cis:
            src_i, bc_method = barcode_matched_cis[ci]
            borrowed_bc, newly = _maybe_backfill(
                cust_barcode_raw, src_i, bc_method, None,
                cust_weight_val=cust_wt, cust_brands_val=cust_brands[ci],
            )
            n_backfilled += int(newly)
            is_bridged = bc_method == 'Barcode (Bridged)'
            blabel = brand_label(src_brand_display[src_i], cust_brand_raw, BRAND_FUZZY_THRESH, alias_dict)

            # ★★ BARCODE-MATCH CONFLICT FLAGGING — a pre-existing barcode
            # link (learned now or in a previous run) is no longer trusted
            # blindly. If weight or brand disagree, surface it as a REVIEW
            # item at reduced confidence instead of silently reporting 100%.
            wm = weights_match(cust_wt, src_weight[src_i])
            weight_conflict = (wm is False)
            brand_conflict = (blabel == 'Mismatch')
            conflict = weight_conflict or brand_conflict

            matched_src_indices.add(src_i)
            if conflict:
                status = 'Matched (Barcode) – REVIEW: possible bad link'
                conf_val = 60
            else:
                status = 'Matched (Barcode-Bridged)' if is_bridged else 'Matched (Barcode)'
                conf_val = 97 if is_bridged else 100

            matched_rows.append({
                **base,
                'Our Brand (brand_standardized)': src_brand_display[src_i],
                'Our Name (parsed)': src_name[src_i],
                'Our Material Code': src_material_code_display[src_i],
                'Our Weight g/ml': src_weight[src_i], 'Our Pack': src_pack[src_i],
                'Our Description (material_name)': src_descraw[src_i],
                'Our Barcode': src_barcode_display[src_i], 'Borrowed Barcode': borrowed_bc,
                'Customer Name (parsed)': cust_name[ci],
                'Match Status': status,
                'Confidence Score': conf_val, 'Method': bc_method,
                'Brand Match': blabel, 'Weight Match': weight_match_label(src_weight[src_i], cust_wt),
            })
            continue

        r = embedded_results.get(ci)
        if r is None or r.get('src_i') is None:
            unmatched_rows.append({**base, 'Reason': (r or {}).get('reason', 'no_result'),
                                   'Best Score': int(round((r or {}).get('score', 0)))})
            continue

        src_i = r['src_i']
        conf  = r['score']
        method = r['method']
        conf_int = int(round(conf))
        borrowed_bc, newly = _maybe_backfill(
            cust_barcode_raw, src_i, method, conf_int,
            cust_weight_val=cust_wt, cust_brands_val=cust_brands[ci],
        )
        n_backfilled += int(newly)
        blabel = brand_label(src_brand_display[src_i], cust_brand_raw, BRAND_FUZZY_THRESH, alias_dict)
        status = 'Matched (High)' if conf_int >= 85 else 'Matched (Medium)'
        matched_src_indices.add(src_i)
        matched_rows.append({
            **base,
            'Our Brand (brand_standardized)': src_brand_display[src_i],
            'Our Name (parsed)': src_name[src_i],
            'Our Material Code': src_material_code_display[src_i],
            'Our Weight g/ml': src_weight[src_i], 'Our Pack': src_pack[src_i],
            'Our Description (material_name)': src_descraw[src_i],
            'Our Barcode': src_barcode_display[src_i], 'Borrowed Barcode': borrowed_bc,
            'Customer Name (parsed)': cust_name[ci],
            'Match Status': status, 'Confidence Score': conf_int, 'Method': method,
            'Brand Match': blabel, 'Weight Match': weight_match_label(src_weight[src_i], cust_wt),
        })

    n_bc  = sum(1 for r in matched_rows if r['Method'] in ('Barcode', 'Barcode (Bridged)'))
    n_emb = sum(1 for r in matched_rows if r['Method'] == 'Embedding')
    n_llm = sum(1 for r in matched_rows if r['Method'] == 'LLM')
    n_review = sum(1 for r in matched_rows if 'REVIEW' in r['Match Status'])
    print(f"  → Matched: {len(matched_rows):,} (BC:{n_bc:,} Emb:{n_emb:,} LLM:{n_llm:,})  "
          f"Unmatched: {len(unmatched_rows):,}  Barcode REVIEW flags: {n_review:,}  "
          f"Barcodes learned: {n_backfilled:,}  ({time.time()-t0:.1f}s)")

    return matched_rows, unmatched_rows, alias_dict, matched_src_indices


# ===========================================================================
# WRITE ALL RESULTS
# ===========================================================================

def write_all_results(all_matched, all_unmatched, alias_dict, per_file, total_src_rows,
                      all_matched_src_indices, src_descraw, cost_tracker, row_cache):
    matched_df   = pd.DataFrame(all_matched)
    unmatched_df = pd.DataFrame(all_unmatched)
    src_covered  = len(all_matched_src_indices)
    src_pct      = src_covered / total_src_rows * 100 if total_src_rows else 0
    total_cust   = len(all_matched) + len(all_unmatched)

    print(f"\nWriting → {OUTPUT_FILE}")
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        matched_df.to_excel(writer, sheet_name='All Matched', index=False)
        if not unmatched_df.empty:
            unmatched_df.to_excel(writer, sheet_name='All Unmatched', index=False)

        # ★ NEW — a dedicated sheet surfacing every barcode match flagged
        # as a possible bad link, across all customers, so these can be
        # spot-checked without hunting through 'All Matched'.
        if not matched_df.empty and 'Match Status' in matched_df.columns:
            review_df = matched_df[matched_df['Match Status'].astype(str).str.contains('REVIEW', na=False)]
            if not review_df.empty:
                review_df.to_excel(writer, sheet_name='Barcode REVIEW Flags', index=False)

        for label, (m_rows, u_rows, _idx) in per_file.items():
            pd.DataFrame(m_rows).to_excel(writer, sheet_name=f"{label} Matched"[:31], index=False)
            if u_rows:
                pd.DataFrame(u_rows).to_excel(writer, sheet_name=f"{label} Unmatched"[:31], index=False)

        rows = [{'Metric': '─── PER CUSTOMER ───', 'Value': '', 'Detail': ''}]
        for label, (m_rows, u_rows, src_idx) in per_file.items():
            t = len(m_rows) + len(u_rows)
            bc  = sum(1 for r in m_rows if r['Method'] in ('Barcode', 'Barcode (Bridged)'))
            emb = sum(1 for r in m_rows if r['Method'] == 'Embedding')
            llm = sum(1 for r in m_rows if r['Method'] == 'LLM')
            review = sum(1 for r in m_rows if 'REVIEW' in r['Match Status'])
            rows += [
                {'Metric': f'[{label}] Rows', 'Value': t, 'Detail': ''},
                {'Metric': f'[{label}] Matched', 'Value': len(m_rows), 'Detail': f"{len(m_rows)/t*100:.1f}%" if t else ''},
                {'Metric': f'[{label}] Unmatched', 'Value': len(u_rows), 'Detail': ''},
                {'Metric': f'[{label}] BC/Emb/LLM', 'Value': f"{bc}/{emb}/{llm}", 'Detail': ''},
                {'Metric': f'[{label}] Barcode REVIEW flags', 'Value': review, 'Detail': ''},
                {'Metric': '', 'Value': '', 'Detail': ''},
            ]
        rows += [
            {'Metric': '─── TOTALS ───', 'Value': '', 'Detail': ''},
            {'Metric': 'Customer rows', 'Value': total_cust, 'Detail': ''},
            {'Metric': 'Matched', 'Value': len(all_matched), 'Detail': ''},
            {'Metric': 'Unmatched', 'Value': len(all_unmatched), 'Detail': ''},
            {'Metric': 'Source coverage', 'Value': src_covered, 'Detail': f"{src_pct:.1f}% of {total_src_rows:,}"},
            {'Metric': '', 'Value': '', 'Detail': ''},
            {'Metric': '─── COST (est.) ───', 'Value': '', 'Detail': ''},
            {'Metric': 'Embed tokens', 'Value': cost_tracker.embed_tokens, 'Detail': f"{cost_tracker.embed_calls} calls"},
            {'Metric': 'LLM in/out tokens', 'Value': f"{cost_tracker.llm_input_tokens}/{cost_tracker.llm_output_tokens}",
             'Detail': f"{cost_tracker.llm_calls} calls"},
            {'Metric': 'Row cache hits', 'Value': cost_tracker.rows_cache_hit, 'Detail': f"of {len(row_cache)} cached"},
            {'Metric': 'LLM hard-rejects (brand/weight veto)', 'Value': cost_tracker.rows_hard_rejected, 'Detail': ''},
            {'Metric': 'Est. USD', 'Value': round(cost_tracker.cost_usd(), 4), 'Detail': ''},
        ]
        pd.DataFrame(rows).to_excel(writer, sheet_name='Summary', index=False)

        never_idx = sorted(set(range(total_src_rows)) - all_matched_src_indices)
        if never_idx:
            pd.DataFrame({'Source Index': never_idx, 'Our Description': [src_descraw[i] for i in never_idx]}
                         ).to_excel(writer, sheet_name='Source Never Matched', index=False)
        pd.DataFrame([{'Source Brand': s, 'Customer Brand': v['Customer Brand'], 'Fuzzy Score': v['Fuzzy Score']}
                      for s, v in sorted(alias_dict.items())]).to_excel(writer, sheet_name='Brand Alias Dict', index=False)

    print(f"  Matched: {len(matched_df):,}  Unmatched: {len(unmatched_df):,}  Coverage: {src_covered:,}/{total_src_rows:,} ({src_pct:.1f}%)")


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    t_total = time.time()
    if not OPENAI_API_KEY:
        print("ERROR: OPENAI_API_KEY not set."); sys.exit(1)

    from openai import OpenAI
    openai_client = OpenAI(
        api_key=OPENAI_API_KEY,
        timeout=API_REQUEST_TIMEOUT_SEC,
        max_retries=0,
        http_client=_make_http_client(),
    )
    cost_tracker  = CostTracker()
    row_cache     = RowMatchCache()

    print(f"Project root: {PROJECT_ROOT}")
    print(f"  ★ Concurrency: LLM_MAX_WORKERS={LLM_MAX_WORKERS}, HTTP_POOL_SIZE={HTTP_POOL_SIZE}")
    if FORCE_RECOMPUTE: print("  ⚠ FORCE_RECOMPUTE = True")
    if SAMPLE_SIZE: print(f"  ⚠ SAMPLE_SIZE = {SAMPLE_SIZE} — this is a TEST run, NOT full data")

    print("\nLoading input …")
    df = load_input_dataframe()
    df = df.dropna(subset=[NAME_COL]).reset_index(drop=True)
    df['source_customer'] = df['source_customer'].astype(str).str.strip()
    bad = df['source_customer'].str.len() > 40
    if bad.any():
        print(f"  WARNING: dropping {int(bad.sum())} corrupt row(s)")
        df = df.loc[~bad].reset_index(drop=True)

    df['_brand_resolved'] = df[BRAND_COL]
    if BRAND_FALLBACK_COL in df.columns:
        blank = df['_brand_resolved'].isna() | (df['_brand_resolved'].astype(str).str.strip() == '')
        df.loc[blank, '_brand_resolved'] = df.loc[blank, BRAND_FALLBACK_COL]
    if CATEGORY_COL not in df.columns:
        print(f"  WARNING: '{CATEGORY_COL}' not found — hierarchy bonus disabled")
        df[CATEGORY_COL] = ''

    src = df[df['source_customer'].str.upper() == SOURCE_CUSTOMER_LABEL.upper()].copy()
    src = src.dropna(subset=[NAME_COL]).drop_duplicates(subset=[NAME_COL]).reset_index(drop=True)
    if src.empty:
        print(f"ERROR: no source rows for '{SOURCE_CUSTOMER_LABEL}'"); sys.exit(1)

    parsed_src     = [parse_description(t) for t in src[NAME_COL]]
    src['_name']   = [p[0] for p in parsed_src]
    src['_weight'] = [p[1] for p in parsed_src]
    src['_kind']   = [p[2] for p in parsed_src]
    src['_pack']   = [p[3] for p in parsed_src]
    src['_brands'] = [normalise_brand_str(b) | brand_tokens_from_desc(nm)
                      for b, nm in zip(src['_brand_resolved'], src['_name'])]

    total_src = len(src)
    print(f"  Source rows ({SOURCE_CUSTOMER_LABEL}): {total_src:,}")

    src_name   = src['_name'].tolist()
    src_weight = src['_weight'].tolist()
    src_kind   = src['_kind'].tolist()
    src_pack   = src['_pack'].tolist()
    src_brands = src['_brands'].tolist()
    src_descraw       = src[NAME_COL].tolist()
    src_brand_display = src['_brand_resolved'].tolist()
    src_h1            = src[CATEGORY_COL].tolist()

    if BARCODE_COL and BARCODE_COL in src.columns:
        src['_barcode'] = src[BARCODE_COL].apply(normalize_barcode)
        src_barcode_display = src[BARCODE_COL].tolist()
    else:
        src['_barcode'] = None
        src_barcode_display = [None] * total_src

    barcode_to_src = {}
    for i, bc in enumerate(src['_barcode'].tolist()):
        if bc and bc not in barcode_to_src: barcode_to_src[bc] = i
    print(f"  Barcode index: {len(barcode_to_src):,} unique")

    src_material_code_display = (src[MATERIAL_CODE_COL].tolist()
                                 if (MATERIAL_CODE_COL and MATERIAL_CODE_COL in src.columns)
                                 else [None] * total_src)

    name_key_to_src = {}
    if BARCODE_COL and BARCODE_COL in df.columns:
        name_to_bcs = defaultdict(set)
        for nm, bc_raw in zip(df[NAME_COL].tolist(), df[BARCODE_COL].tolist()):
            key, bc = normalize_name_key(nm), normalize_barcode(bc_raw)
            if key and bc: name_to_bcs[key].add(bc)
        for key, bcs in name_to_bcs.items():
            hits = {barcode_to_src[bc] for bc in bcs if bc in barcode_to_src}
            if len(hits) == 1: name_key_to_src[key] = next(iter(hits))
        print(f"  Barcode↔name bridge: {len(name_key_to_src):,} entries")

    print("\nBuilding FAISS index …")
    faiss_index = OpenAIFaissIndex(src_name, cost_tracker)

    customer_labels = sorted(
        df.loc[df['source_customer'].str.upper() != SOURCE_CUSTOMER_LABEL.upper(), 'source_customer'].unique())
    print(f"\nCustomers: {customer_labels}")

    all_matched, all_unmatched, all_aliases = [], [], {}
    per_file, all_matched_src_indices = {}, set()

    try:
        for label in customer_labels:
            cust = get_customer_frame(df, label, barcode_to_src=barcode_to_src)
            if cust.empty: continue

            m_rows, u_rows, aliases, src_idx_set = match_customer(
                label, cust,
                src_name, src_weight, src_kind, src_pack, src_brands, src_h1,
                src_descraw, src_brand_display, src_barcode_display, src_material_code_display,
                barcode_to_src, name_key_to_src,
                faiss_index, openai_client, cost_tracker, row_cache,
            )

            all_matched.extend(m_rows)
            all_unmatched.extend(u_rows)
            all_aliases.update(aliases)
            per_file[label] = (m_rows, u_rows, src_idx_set)
            all_matched_src_indices |= src_idx_set

            if cost_tracker.budget_exceeded():
                print(f"\n  ⚠ BUDGET reached — stopping.")
                break
    finally:
        cost_tracker.stop()
        row_cache.flush()
        print(f"  Row cache saved: {len(row_cache):,} entries → {ROW_CACHE_FILE}")

    write_all_results(all_matched, all_unmatched, all_aliases, per_file, total_src,
                      all_matched_src_indices, src_descraw, cost_tracker, row_cache)

    print(f"\nTotal time: {time.time()-t_total:.1f}s")
    print(f"Saved → {OUTPUT_FILE}")


if __name__ == '__main__':
    main()