"""
MatchIQ v2 — Fixed Weight Logic
================================

Weight logic fix summary:
  OLD (broken):  both must have weight OR both must be missing → too strict
  NEW (fixed):
    • If BOTH have a weight  → they must match (strict check)
    • If ONLY customer has   → still allow match, flag as "Weight Only Customer"
    • If ONLY source has     → still allow match, flag as "Weight Only Source"
    • If NEITHER has weight  → allow match, flag as "N/A"

  Weight matching now SCORES candidates instead of hard-rejecting them:
    Priority 1 → both have weight AND weights match   (best)
    Priority 2 → neither has weight
    Priority 3 → only one side has weight             (allowed but flagged)
    Priority 4 → both have weight but they differ     (only in soft mode)

Run:  python matchiq_v2_fixed.py
"""

import re
import sys
import time
import logging
import warnings
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
from tqdm import tqdm
from rapidfuzz import fuzz, process as rf_process
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

warnings.filterwarnings("ignore")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("matchiq")


# ─────────────────────────────────────────────────────────────────────────────
# ██  CONFIG — EDIT THIS SECTION  ██
# ─────────────────────────────────────────────────────────────────────────────

CONFIG = {
    "source": {
        "file": r"C:\Users\HP\Desktop\functionand sp.csv",
        "column": "material_desc",
    },

    "targets": [
        {
            "file": r"C:\Users\HP\Downloads\fwqniemtselloutdashboardsession2\C4 sep 24.xlsx",
            "column": "Item Name",
        },
        # {"file": r"C:\path\to\customer2.csv", "column": "Product Name"},
    ],

    "output": r"C:\Users\HP\Downloads\matchiq_v2_fixed_results.xlsx",

    # 100 for quick test, None for full file
    "sample": 300,

    "thresholds": {
        "fuzzy":    85,
        "tfidf":    80,
        "semantic": 80,
    },

    "layers": {
        "fuzzy":    True,
        "tfidf":    True,
        "semantic": True,   # pip install sentence-transformers
    },

    "top_k": {
        "fuzzy":    10,
        "tfidf":    15,
        "semantic": 15,
    },

    # "strict" → reject candidates where BOTH sides have weights but they differ
    # "soft"   → allow weight mismatch but flag it
    "weight_rule": "soft",   # ← changed to soft so more matches come through

    "performance": {
        "semantic_model": "all-MiniLM-L6-v2",
        "semantic_batch": 128,
        "chunk_size":     500,
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# END OF CONFIG
# ─────────────────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────────────────
# LAYER 1 — NORMALIZATION + PARSING
# ─────────────────────────────────────────────────────────────────────────────

UNIT_MAP = {
    r"\blitre[s]?\b": "l",       r"\bliter[s]?\b": "l",    r"\bltr[s]?\b": "l",
    r"\bkilogram[s]?\b": "kg",   r"\bkilos?\b": "kg",       r"\bkgs?\b": "kg",
    r"\bgram[s]?\b": "g",        r"\bgrm[s]?\b": "g", r"\bgr?\b": "g",  r"\bgm[s]?\b": "g", r"\bG\b": "g",
    r"\bmillilitre[s]?\b": "ml", r"\bmilliliter[s]?\b": "ml",
    r"\bpound[s]?\b": "lb",      r"\bounce[s]?\b": "oz",
}
_COMPILED_UNIT_MAP = [(re.compile(p), r) for p, r in UNIT_MAP.items()]

STOPWORDS = frozenset({
    "of", "the", "a", "an", "and", "or", "with", "for",
    "in", "on", "at", "to", "by", "new", "fresh",
})

_RE_SPECIAL    = re.compile(r"[^a-z0-9\s]")
_RE_WHITESPACE = re.compile(r"\s+")

WEIGHT_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*"
    r"(g|gm|gms|gr|grm|grms|gram|grams|kg|ml|l|ltr|litre|liter|oz|cl)\b",
    re.IGNORECASE,
)
PACK_RE = re.compile(r"(\d+\s*x\s*(?:\d+\s*x\s*)*)", re.IGNORECASE)


def parse_description(text: str) -> tuple:
    """Return (name_string, pack_string, weight_string)."""
    if not isinstance(text, str) or not text.strip():
        return "", "", ""
    s = text.strip().lower()

    # grab the LAST weight token (unit weight, not pack weight)
    weight = ""
    wmatches = list(WEIGHT_RE.finditer(s))
    if wmatches:
        wm = wmatches[-1]
        weight = wm.group(0).replace(" ", "")
        s = s[:wm.start()] + " " + s[wm.end():]

    # grab pack pattern  e.g. 4x, 12x6, 2x3
    pack = ""
    pm = PACK_RE.search(s)
    if pm:
        pack = pm.group(0).replace(" ", "").rstrip("x").lower()
        s = s[:pm.start()] + " " + s[pm.end():]

    name = re.sub(r"\s+", " ", s).strip()
    return name, pack, weight


def normalize_name(name: str) -> str:
    if not name:
        return ""
    t = name.lower()
    for pat, rep in _COMPILED_UNIT_MAP:
        t = pat.sub(rep, t)
    t = _RE_SPECIAL.sub(" ", t)
    t = _RE_WHITESPACE.sub(" ", t).strip()
    return " ".join(w for w in t.split() if w not in STOPWORDS and len(w) > 1)


def normalize_weight(w: str) -> Optional[float]:
    """Normalise to a common unit: grams for mass, ml for volume."""
    if not w:
        return None
    m = re.match(
        r"(\d+(?:\.\d+)?)(g|gm|gms|gr|grm|grms|gram|grams|kg|ml|l|ltr|litre|liter|oz|cl)",
        w.strip(), re.IGNORECASE,
    )
    if not m:
        return None
    val  = float(m.group(1))
    unit = m.group(2).lower()
    if unit in ("g", "gm", "gms", "gram", "grams"): return val
    if unit == "kg":                                  return val * 1000
    if unit == "ml":                                  return val
    if unit in ("l", "ltr", "litre", "liter"):        return val * 1000
    if unit == "cl":                                  return val * 10
    if unit == "oz":                                  return round(val * 28.3495, 2)
    return None


def normalize_pack(p: str) -> Optional[str]:
    if not p:
        return None
    c = p.strip().lower().rstrip("x")
    return c or None


# ─────────────────────────────────────────────────────────────────────────────
# WEIGHT / PACK SCORING  ← FIXED
# ─────────────────────────────────────────────────────────────────────────────

# Weight priority levels (lower = better)
# 0 = both have weight and they match          ← best
# 1 = neither side has a weight                ← good
# 2 = only one side has a weight               ← acceptable, flagged
# 3 = both have weight but they DON'T match    ← worst (blocked in strict mode)

def weight_priority(src_w: Optional[float], cust_w: Optional[float]) -> int:
    """Return a priority score for this weight pair (0 = best, 3 = worst)."""
    if src_w is not None and cust_w is not None:
        return 0 if src_w == cust_w else 3   # both present: must match
    if src_w is None and cust_w is None:
        return 1                              # neither has weight
    return 2                                  # only one side has weight


def pack_matches(src_p: Optional[str], cust_p: Optional[str]) -> bool:
    if src_p is None and cust_p is None:
        return True
    if src_p is not None and cust_p is not None:
        return src_p == cust_p
    return False                              # one side missing → no pack match


def weight_match_label(src_w: Optional[float], cust_w: Optional[float]) -> str:
    """Human-readable label for the output report."""
    if src_w is None and cust_w is None:
        return "N/A – no weight on either side"
    if src_w is not None and cust_w is not None:
        if src_w == cust_w:
            return f"Matched – both {src_w}{'g' if src_w < 10000 else 'ml'}"
        else:
            return f"Mismatch – source {src_w} vs customer {cust_w}"
    if src_w is not None:
        return f"Source only – {src_w}"
    return f"Customer only – {cust_w}"


# ─────────────────────────────────────────────────────────────────────────────
# CANDIDATE SELECTOR  ← FIXED
# ─────────────────────────────────────────────────────────────────────────────

def select_best_candidate(
    candidates: list,            # [(source_idx, text_score), ...]
    cust_w:     Optional[float],
    cust_p:     Optional[str],
    source_parsed: list,
    weight_rule:   str = "soft",
) -> tuple:
    """
    Rank candidates by:
      1. weight_priority  (0 best → 3 worst)
      2. pack match bonus
      3. text score (higher = better)

    In STRICT mode: candidates with weight_priority == 3
                    (both have weights but they differ) are dropped entirely.
    In SOFT mode  : they stay but rank last.

    Returns (best_idx, best_score, weight_label) or (None, 0.0, "Unmatched").
    """
    if not candidates:
        return None, 0.0, "Unmatched"

    scored = []
    for idx, text_score in candidates:
        sp   = source_parsed[idx]
        wpri = weight_priority(sp["weight_n"], cust_w)

        # strict mode: skip candidates where both sides have weights but differ
        if weight_rule == "strict" and wpri == 3:
            continue

        pack_bonus = 1 if pack_matches(sp["pack"], cust_p) else 0

        # sort key: (weight_priority ASC, pack_bonus DESC, text_score DESC)
        scored.append((wpri, -pack_bonus, -text_score, idx, text_score))

    if not scored:
        return None, 0.0, "Unmatched"

    scored.sort()
    _, _, _, best_idx, best_score = scored[0]
    sp    = source_parsed[best_idx]
    label = weight_match_label(sp["weight_n"], cust_w)
    return best_idx, best_score, label


# ─────────────────────────────────────────────────────────────────────────────
# LAYER 2 — FUZZY (top-K)
# ─────────────────────────────────────────────────────────────────────────────

def fuzzy_top_k(target_name: str, source_names: list, k: int, threshold: int) -> list:
    if not target_name:
        return []
    hits = rf_process.extract(
        target_name, source_names,
        scorer=fuzz.token_sort_ratio,
        limit=k, score_cutoff=threshold,
    )
    return [(idx, float(score)) for _, score, idx in hits]


# ─────────────────────────────────────────────────────────────────────────────
# LAYER 3 — TF-IDF (top-K)
# ─────────────────────────────────────────────────────────────────────────────

class TFIDFMatcher:
    def __init__(self, source_names: list):
        log.info("Building TF-IDF index on %d source names …", len(source_names))
        self.vectorizer = TfidfVectorizer(
            analyzer="char_wb",
            ngram_range=(3, 5),
            min_df=1,
            sublinear_tf=True,
        )
        self.matrix = self.vectorizer.fit_transform(source_names)

    def top_k(self, target_name: str, k: int, threshold: float) -> list:
        if not target_name:
            return []
        v    = self.vectorizer.transform([target_name])
        sims = cosine_similarity(v, self.matrix)[0]
        return [
            (int(i), round(float(sims[i]) * 100, 1))
            for i in np.argsort(sims)[::-1][:k]
            if float(sims[i]) >= threshold
        ]


# ─────────────────────────────────────────────────────────────────────────────
# LAYER 4 — SEMANTIC (top-K)
# ─────────────────────────────────────────────────────────────────────────────

class SemanticMatcher:
    def __init__(self, source_names: list, model_name: str, batch_size: int):
        log.info("Loading embedding model '%s' …", model_name)
        try:
            from sentence_transformers import SentenceTransformer
        except ImportError:
            raise ImportError("Run: pip install sentence-transformers")
        self.model      = SentenceTransformer(model_name)
        self.batch_size = batch_size
        log.info("Encoding %d source names …", len(source_names))
        self.embeddings = self.model.encode(
            source_names,
            batch_size=batch_size,
            show_progress_bar=True,
            convert_to_numpy=True,
            normalize_embeddings=True,
        )

    def encode(self, names: list, batch_size: int) -> np.ndarray:
        return self.model.encode(
            names,
            batch_size=batch_size,
            show_progress_bar=False,
            convert_to_numpy=True,
            normalize_embeddings=True,
        )

    def top_k_for_embedding(self, t_emb: np.ndarray, k: int, threshold: float) -> list:
        sims = self.embeddings @ t_emb
        return [
            (int(i), round(float(sims[i]) * 100, 1))
            for i in np.argsort(sims)[::-1][:k]
            if float(sims[i]) >= threshold
        ]


# ─────────────────────────────────────────────────────────────────────────────
# FILE I/O
# ─────────────────────────────────────────────────────────────────────────────

def _find_column_in_rows(df_raw: pd.DataFrame, col_name: str) -> int:
    target = col_name.strip().lower()
    for row_idx in range(min(30, len(df_raw))):
        if target in df_raw.iloc[row_idx].astype(str).str.strip().str.lower().values:
            return row_idx
    return -1


def _make_df_from_header_row(df_raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    df          = df_raw.copy()
    df.columns  = df.iloc[header_row].astype(str).str.strip()
    df          = df.iloc[header_row + 1:].reset_index(drop=True)
    df.dropna(how="all", inplace=True)
    df.columns  = [str(c).strip() for c in df.columns]
    return df


def read_file(path, col):
    path = Path(path)
    log.info("Reading '%s' — column '%s' …", path.name, col)
    ext  = path.suffix.lower()

    if ext == ".csv":
        for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
            try:
                df_raw = pd.read_csv(path, header=None, dtype=str, encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError(f"Cannot decode CSV '{path}'.")
        hr = _find_column_in_rows(df_raw, col)
        if hr < 0:
            raise ValueError(f"Column '{col}' not found in '{path.name}'.")
        df = _make_df_from_header_row(df_raw, hr)

    elif ext in {".xlsx", ".xls", ".xlsb"}:
        engine = "pyxlsb" if ext == ".xlsb" else "openpyxl"
        try:
            xls = pd.ExcelFile(path, engine=engine)
        except Exception:
            xls = pd.ExcelFile(path, engine="xlrd")
        found_raw, found_hr = None, -1
        for sname in xls.sheet_names:
            try:
                raw = pd.read_excel(xls, sheet_name=sname, header=None, dtype=str)
                hr  = _find_column_in_rows(raw, col)
                if hr >= 0:
                    found_raw, found_hr = raw, hr
                    log.info("  → sheet '%s', header row %d", sname, hr)
                    break
            except Exception as e:
                log.warning("  sheet '%s' skipped: %s", sname, e)
        if found_raw is None:
            raise ValueError(f"Column '{col}' not found in any sheet of '{path.name}'.")
        df = _make_df_from_header_row(found_raw, found_hr)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    df          = df.dropna(how="all").reset_index(drop=True)
    df.columns  = [str(c).strip() for c in df.columns]
    if col not in df.columns:
        raise ValueError(f"Column '{col}' missing after cleanup. Got: {list(df.columns)}")
    log.info("  → %d rows loaded", len(df))
    df["__desc_col__"] = df[col].fillna("").astype(str).str.strip()
    return df, col


def write_results(results: pd.DataFrame, output_path):
    output_path = Path(output_path)
    log.info("Writing '%s' …", output_path)
    total = len(results)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Matched rows only
        matched_df = results[results["Match Status"] == "Matched"]
        matched_df.to_excel(writer, sheet_name="Matched Results", index=False)

        # Method summary
        pd.DataFrame([
            {
                "Method": m,
                "Count":  (cnt := (results["Method"] == m).sum()),
                "Pct":    f"{round(cnt/total*100,1)}%" if total else "0%",
            }
            for m in ["Fuzzy", "TF-IDF", "Semantic", "None"]
        ]).to_excel(writer, sheet_name="Summary", index=False)

        # Weight summary — now uses the richer labels
        w_counts = results["Weight Match"].value_counts().reset_index()
        w_counts.columns = ["Weight Match", "Count"]
        w_counts["Pct"] = (w_counts["Count"] / total * 100).round(1).astype(str) + "%"
        w_counts.to_excel(writer, sheet_name="Weight Summary", index=False)

        # Unmatched rows
        un = results[results["Match Status"] == "Unmatched"]
        if not un.empty:
            un.to_excel(writer, sheet_name="Unmatched", index=False)

    log.info("Done. %d rows written.", total)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def run_pipeline(cfg: dict):
    t0          = time.time()
    th          = cfg["thresholds"]
    layers      = cfg["layers"]
    topk        = cfg["top_k"]
    perf        = cfg["performance"]
    weight_rule = cfg.get("weight_rule", "soft")
    sample_n    = cfg.get("sample")

    fuzzy_th = th["fuzzy"]
    tfidf_th = th["tfidf"] / 100
    sem_th   = th["semantic"] / 100

    # ── Source ──────────────────────────────────────────────────────────────
    src_df, _ = read_file(cfg["source"]["file"], cfg["source"]["column"])
    source_raw = src_df["__desc_col__"].tolist()

    log.info("Parsing %d source descriptions …", len(source_raw))
    source_parsed = []
    for raw in source_raw:
        name, pack, weight = parse_description(raw)
        source_parsed.append({
            "raw":        raw,
            "name_norm":  normalize_name(name),
            "pack":       normalize_pack(pack),
            "weight_raw": weight,
            "weight_n":   normalize_weight(weight),
        })
    source_names = [p["name_norm"] for p in source_parsed]

    # debug: show a few parsed examples so you can verify weight extraction
    log.info("─── Source parse sample (first 5) ───")
    for p in source_parsed[:5]:
        log.info("  raw=%-40s  weight_raw=%-8s  weight_n=%s",
                 p["raw"][:40], p["weight_raw"], p["weight_n"])
    log.info("─────────────────────────────────────")

    # ── Build indexes ─────────────────────────────────────────────────────
    tfidf    = TFIDFMatcher(source_names) if layers["tfidf"] else None
    semantic = (
        SemanticMatcher(source_names, perf["semantic_model"], perf["semantic_batch"])
        if layers["semantic"] else None
    )

    # ── Process targets ────────────────────────────────────────────────────
    all_results = []

    for tgt_cfg in cfg["targets"]:
        tgt_df, _ = read_file(tgt_cfg["file"], tgt_cfg["column"])

        if sample_n is not None:
            tgt_df = tgt_df.head(sample_n)
            log.info("  → SAMPLE MODE: %d rows (set 'sample': None for full run)", sample_n)

        target_raw = tgt_df["__desc_col__"].tolist()
        file_name  = Path(tgt_cfg["file"]).name

        log.info("Parsing %d target descriptions …", len(target_raw))
        target_parsed = []
        for raw in target_raw:
            name, pack, weight = parse_description(raw)
            target_parsed.append({
                "raw":        raw,
                "name_norm":  normalize_name(name),
                "pack":       normalize_pack(pack),
                "weight_raw": weight,
                "weight_n":   normalize_weight(weight),
            })

        # debug: show a few parsed target examples
        log.info("─── Target parse sample (first 5) ───")
        for p in target_parsed[:5]:
            log.info("  raw=%-40s  weight_raw=%-8s  weight_n=%s",
                     p["raw"][:40], p["weight_raw"], p["weight_n"])
        log.info("─────────────────────────────────────")

        # pre-encode semantic embeddings for all targets at once
        target_emb = None
        if semantic is not None:
            target_emb = semantic.encode(
                [p["name_norm"] for p in target_parsed],
                batch_size=perf["semantic_batch"],
            )

        rows = []
        for i, tp in enumerate(tqdm(target_parsed,
                                    desc=f"Matching {file_name[:30]}",
                                    unit="row")):
            cust_name = tp["name_norm"]
            cust_w    = tp["weight_n"]
            cust_p    = tp["pack"]

            best = {"idx": None, "score": 0.0, "method": "None", "wlabel": "Unmatched"}

            # ── Layer 2: Fuzzy ────────────────────────────────────────────
            if layers["fuzzy"] and cust_name:
                cands = fuzzy_top_k(cust_name, source_names, topk["fuzzy"], fuzzy_th)
                idx, sc, wlabel = select_best_candidate(
                    cands, cust_w, cust_p, source_parsed, weight_rule)
                if idx is not None:
                    best = {"idx": idx, "score": sc, "method": "Fuzzy", "wlabel": wlabel}

            # ── Layer 3: TF-IDF ───────────────────────────────────────────
            if best["idx"] is None and tfidf and cust_name:
                cands = tfidf.top_k(cust_name, topk["tfidf"], tfidf_th)
                idx, sc, wlabel = select_best_candidate(
                    cands, cust_w, cust_p, source_parsed, weight_rule)
                if idx is not None:
                    best = {"idx": idx, "score": sc, "method": "TF-IDF", "wlabel": wlabel}

            # ── Layer 4: Semantic ─────────────────────────────────────────
            if best["idx"] is None and semantic and cust_name:
                cands = semantic.top_k_for_embedding(
                    target_emb[i], topk["semantic"], sem_th)
                idx, sc, wlabel = select_best_candidate(
                    cands, cust_w, cust_p, source_parsed, weight_rule)
                if idx is not None:
                    best = {"idx": idx, "score": sc, "method": "Semantic", "wlabel": wlabel}

            # ── Build output row ──────────────────────────────────────────
            if best["idx"] is not None:
                sp = source_parsed[best["idx"]]
                rows.append({
                    "Source File":              file_name,
                    "Our Description":          sp["raw"],
                    "Our Weight (parsed)":      sp["weight_raw"] or "—",
                    "Customer Description":     tp["raw"],
                    "Customer Weight (parsed)": tp["weight_raw"] or "—",
                    "Match Status":             "Matched",
                    "Score":                    best["score"],
                    "Method":                   best["method"],
                    "Weight Match":             best["wlabel"],
                })
            else:
                rows.append({
                    "Source File":              file_name,
                    "Our Description":           sp["raw"],
                    "Our Weight (parsed)":       sp["weight_raw"] or "—",
                    "Customer Description":     tp["raw"],
                    "Customer Weight (parsed)": tp["weight_raw"] or "—",
                    "Match Status":             "Unmatched",
                    "Score":                    0.0,
                    "Method":                   "None",
                    "Weight Match":             "Unmatched",
                })

        result_df = pd.DataFrame(rows)
        all_results.append(result_df)

        matched = (result_df["Match Status"] == "Matched").sum()
        log.info("  '%s' → %d/%d matched (%.1f%%)",
                 file_name, matched, len(result_df),
                 matched / len(result_df) * 100 if len(result_df) else 0)

    final         = pd.concat(all_results, ignore_index=True)
    total         = len(final)
    matched_total = (final["Match Status"] == "Matched").sum()

    log.info("=" * 60)
    log.info("SUMMARY")
    log.info("  Total       : %d", total)
    log.info("  Matched     : %d (%.1f%%)",
             matched_total, matched_total / total * 100 if total else 0)
    for m in ["Fuzzy", "TF-IDF", "Semantic"]:
        log.info("    %-9s: %d", m, (final["Method"] == m).sum())
    wt_counts = final["Weight Match"].value_counts()
    for label, cnt in wt_counts.items():
        log.info("  Weight %-40s : %d", label, cnt)
    log.info("  Time        : %.1fs", time.time() - t0)
    log.info("=" * 60)

    write_results(final, cfg["output"])
    return final


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY
# ─────────────────────────────────────────────────────────────────────────────

def main():
    cfg = CONFIG
    if not Path(cfg["source"]["file"]).exists():
        log.error("Source file not found: %s", cfg["source"]["file"])
        sys.exit(1)
    missing = [t["file"] for t in cfg["targets"] if not Path(t["file"]).exists()]
    if missing:
        log.error("Target file(s) not found: %s", missing)
        sys.exit(1)

    log.info("MatchIQ v2 (fixed weight logic) starting …")
    log.info("  Layers      : fuzzy=%s  tfidf=%s  semantic=%s",
             cfg["layers"]["fuzzy"], cfg["layers"]["tfidf"], cfg["layers"]["semantic"])
    log.info("  Weight rule : %s", cfg.get("weight_rule", "soft"))
    log.info("  Sample      : %s",
             f"{cfg['sample']} rows" if cfg.get("sample") else "FULL FILE")
    run_pipeline(cfg)


if __name__ == "__main__":
    main()