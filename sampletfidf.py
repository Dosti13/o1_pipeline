"""
MatchIQ v2 — Smart Description Matcher (weight & pack aware)
=============================================================

Combines:
  • The 4-layer pipeline (Fuzzy → TF-IDF → Semantic) from matchiq.py
  • The name / pack / weight parsing logic from your fuzzy_match v2 script

For EVERY layer we now:
  1. Get the top-K candidates by text similarity
  2. Filter / re-rank by WEIGHT (must match if both sides have one)
  3. Among weight-matched, prefer PACK match
  4. Pick the highest-scoring survivor

Output columns (per row):
    Source File
    Our Description           ← matched source description
    Customer Description      ← raw target description
    Match Status              ← Matched / Unmatched
    Score                     ← 0–100 confidence
    Method                    ← Fuzzy / TF-IDF / Semantic / None
    Weight Match              ← With Weight / Without Weight / N/A

Run:  python matchiq_v2.py
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
        {"file": r"C:\Users\HP\Downloads\fwqniemtselloutdashboardsession2\C4 sep 24.xlsx",
         "column": "Item Name"},
        # add more target files as needed:
        # {"file": r"C:\path\to\customer2.csv", "column": "Product Name"},
    ],

    "output": r"C:\Users\HP\Downloads\matchiq_v2_results.xlsx",

    # ── Sample limit ──────────────────────────────────────────────────────
    # Set to 100 for a quick test run. Change to None for the full file.
    "sample": None,

    # ── Thresholds (0–100, higher = stricter) ─────────────────────────────
    "thresholds": {
        "fuzzy": 80,
        "tfidf": 80,
        "semantic": 80,
    },

    # ── Layer toggles ─────────────────────────────────────────────────────
    "layers": {
        "fuzzy": True,
        "tfidf": True,
        "semantic": True,   # set True after `pip install sentence-transformers`
    },

    # ── How many top candidates each layer hands to the weight/pack judge ─
    "top_k": {
        "fuzzy": 10,
        "tfidf": 15,
        "semantic": 15,
    },

    # ── Weight rule ───────────────────────────────────────────────────────
    # "strict": if customer has a weight, source MUST have the same weight.
    # "soft":   weights matter, but mismatches are still allowed (just flagged).
    "weight_rule": "strict",

    "performance": {
        "semantic_model": "all-MiniLM-L6-v2",
        "semantic_batch": 128,
        "chunk_size": 500,
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# END OF CONFIG
# ─────────────────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────────────────
# LAYER 1 — NORMALIZATION + (NAME / PACK / WEIGHT) PARSING
# ─────────────────────────────────────────────────────────────────────────────

UNIT_MAP = {
    r"\blitre[s]?\b": "l",
    r"\bliter[s]?\b": "l",
    r"\bltr[s]?\b": "l",
    r"\bkilogram[s]?\b": "kg",
    r"\bkilos?\b": "kg",
    r"\bkgs?\b": "kg",
    r"\bG\b": "g",
    r"\bgm[s]?\b": "g",
    r"\bgr?\b": "g",
    r"\bgrm[s]?\b": "g",
    r"\bgram[s]?\b": "g",
    r"\bgrm[s]?\b": "g",
    r"\bmillilitre[s]?\b": "ml",
    r"\bmilliliter[s]?\b": "ml",
    r"\bpound[s]?\b": "lb",
    r"\bounce[s]?\b": "oz",
}
_COMPILED_UNIT_MAP = [(re.compile(p), r) for p, r in UNIT_MAP.items()]

STOPWORDS = frozenset({
    "of", "the", "a", "an", "and", "or", "with", "for",
    "in", "on", "at", "to", "by", "new", "fresh",
})

_RE_SPECIAL = re.compile(r"[^a-z0-9\s]")
_RE_WHITESPACE = re.compile(r"\s+")

WEIGHT_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*(g|gm|gms|gram|grams|kg|ml|l|ltr|litre|liter|oz|cl)\b",
    re.IGNORECASE,
)
PACK_RE = re.compile(r"(\d+\s*x\s*(?:\d+\s*x\s*)*)", re.IGNORECASE)


def parse_description(text: str) -> tuple:
    """Split a product description into (name, pack, weight)."""
    if not isinstance(text, str) or not text.strip():
        return "", "", ""
    s = text.strip().lower()

    weight = ""
    weight_matches = list(WEIGHT_RE.finditer(s))
    if weight_matches:
        wm = weight_matches[-1]
        weight = wm.group(0).replace(" ", "")
        s = s[: wm.start()] + " " + s[wm.end():]

    pack = ""
    pack_match = PACK_RE.search(s)
    if pack_match:
        pack = pack_match.group(0).replace(" ", "").rstrip("x").lower()
        s = s[: pack_match.start()] + " " + s[pack_match.end():]

    name = re.sub(r"\s+", " ", s).strip()
    return name, pack, weight


def normalize_name(name: str) -> str:
    """Clean a product NAME (after weight & pack are stripped) for matching."""
    if not name:
        return ""
    t = name.lower()
    for pattern, replacement in _COMPILED_UNIT_MAP:
        t = pattern.sub(replacement, t)
    t = _RE_SPECIAL.sub(" ", t)
    t = _RE_WHITESPACE.sub(" ", t).strip()
    tokens = [w for w in t.split() if w not in STOPWORDS and len(w) > 1]
    return " ".join(tokens)


def normalize_weight(w: str) -> Optional[float]:
    """Convert any weight string to grams (mass) or ml (volume)."""
    if not w:
        return None
    m = re.match(
        r"(\d+(?:\.\d+)?)(g|gm|gms|gram|grams|kg|ml|l|ltr|litre|liter|oz|cl)",
        w, re.IGNORECASE,
    )
    if not m:
        return None
    val = float(m.group(1))
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
    clean = p.strip().lower().rstrip("x")
    return clean if clean else None


# ─────────────────────────────────────────────────────────────────────────────
# WEIGHT / PACK COMPATIBILITY
# ─────────────────────────────────────────────────────────────────────────────

def weight_compatible(our_w: Optional[float], cust_w: Optional[float]) -> bool:
    if our_w is None and cust_w is None:
        return True
    if our_w is not None and cust_w is not None:
        return our_w == cust_w
    return False


def pack_compatible(our_p: Optional[str], cust_p: Optional[str]) -> bool:
    if our_p is None and cust_p is None:
        return True
    if our_p is not None and cust_p is not None:
        return our_p == cust_p
    return False


def weight_match_label(our_w, cust_w) -> str:
    if our_w is None and cust_w is None:
        return "N/A (no weight on either side)"
    if weight_compatible(our_w, cust_w):
        return "With Weight"
    return "Without Weight"


# ─────────────────────────────────────────────────────────────────────────────
# LAYER 2 — FUZZY (top-K)
# ─────────────────────────────────────────────────────────────────────────────

def fuzzy_top_k(target_name: str, source_names: list, k: int, threshold: int):
    if not target_name:
        return []
    hits = rf_process.extract(
        target_name,
        source_names,
        scorer=fuzz.token_sort_ratio,
        limit=k,
        score_cutoff=threshold,
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

    def top_k(self, target_name: str, k: int, threshold: float):
        if not target_name:
            return []
        v = self.vectorizer.transform([target_name])
        sims = cosine_similarity(v, self.matrix)[0]
        idx = np.argsort(sims)[::-1][:k]
        out = []
        for i in idx:
            sc = float(sims[i])
            if sc >= threshold:
                out.append((int(i), round(sc * 100, 1)))
        return out


# ─────────────────────────────────────────────────────────────────────────────
# LAYER 4 — SEMANTIC (top-K)
# ─────────────────────────────────────────────────────────────────────────────

class SemanticMatcher:
    def __init__(self, source_names: list, model_name: str, batch_size: int):
        log.info("Loading embedding model '%s' …", model_name)
        try:
            from sentence_transformers import SentenceTransformer
        except ImportError:
            raise ImportError(
                "sentence-transformers not installed.\n"
                "Run: pip install sentence-transformers"
            )
        self.model = SentenceTransformer(model_name)
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

    def top_k_for_embedding(self, t_emb: np.ndarray, k: int, threshold: float):
        sims = self.embeddings @ t_emb
        idx = np.argsort(sims)[::-1][:k]
        out = []
        for i in idx:
            sc = float(sims[i])
            if sc >= threshold:
                out.append((int(i), round(sc * 100, 1)))
        return out


# ─────────────────────────────────────────────────────────────────────────────
# CANDIDATE SELECTOR — applies the Name → Weight → Pack priority
# ─────────────────────────────────────────────────────────────────────────────

def select_best_candidate(
    candidates: list,
    cust_w: Optional[float],
    cust_p: Optional[str],
    source_parsed: list,
    weight_rule: str = "strict",
):
    if not candidates:
        return None, 0.0, False

    weight_ok = []
    weight_bad = []
    for idx, score in candidates:
        sp = source_parsed[idx]
        if weight_compatible(sp["weight_n"], cust_w):
            weight_ok.append((idx, score))
        else:
            weight_bad.append((idx, score))

    if weight_ok:
        pack_ok = [c for c in weight_ok
                   if pack_compatible(source_parsed[c[0]]["pack"], cust_p)]
        pool = pack_ok if pack_ok else weight_ok
        best_idx, best_score = max(pool, key=lambda x: x[1])
        return best_idx, best_score, True

    if weight_rule == "strict":
        return None, 0.0, False

    best_idx, best_score = max(weight_bad, key=lambda x: x[1])
    return best_idx, best_score, False


# ─────────────────────────────────────────────────────────────────────────────
# FILE I/O
# ─────────────────────────────────────────────────────────────────────────────

def _find_column_in_rows(df_raw: pd.DataFrame, col_name: str) -> int:
    target = col_name.strip().lower()
    for row_idx in range(min(30, len(df_raw))):
        row_vals = df_raw.iloc[row_idx].astype(str).str.strip().str.lower()
        if target in row_vals.values:
            return row_idx
    return -1


def _make_df_from_header_row(df_raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    df = df_raw.copy()
    df.columns = df.iloc[header_row].astype(str).str.strip()
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.dropna(how="all", inplace=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def read_file(path, col):
    path = Path(path)
    log.info("Reading '%s' — looking for column '%s' …", path.name, col)
    ext = path.suffix.lower()

    if ext == ".csv":
        for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
            try:
                df_raw = pd.read_csv(path, header=None, dtype=str, encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError(f"Cannot decode CSV '{path}'.")
        header_row = _find_column_in_rows(df_raw, col)
        if header_row == -1:
            raise ValueError(f"Column '{col}' not found in '{path.name}'.")
        log.info("  → found at header row %d", header_row)
        df = _make_df_from_header_row(df_raw, header_row)

    elif ext in {".xlsx", ".xls", ".xlsb"}:
        engine = "pyxlsb" if ext == ".xlsb" else "openpyxl"
        try:
            xls = pd.ExcelFile(path, engine=engine)
        except Exception:
            xls = pd.ExcelFile(path, engine="xlrd")

        found_sheet = found_df_raw = None
        found_header_row = -1
        for sname in xls.sheet_names:
            try:
                df_raw = pd.read_excel(xls, sheet_name=sname, header=None, dtype=str)
                hr = _find_column_in_rows(df_raw, col)
                if hr >= 0:
                    found_sheet, found_df_raw, found_header_row = sname, df_raw, hr
                    log.info("  → found in sheet '%s' at row %d", sname, hr)
                    break
            except Exception as e:
                log.warning("  → sheet '%s' skipped: %s", sname, e)

        if found_sheet is None:
            raise ValueError(f"Column '{col}' not found in any sheet of '{path.name}'.")
        df = _make_df_from_header_row(found_df_raw, found_header_row)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    if col not in df.columns:
        raise ValueError(f"Column '{col}' lost after cleanup. Got: {list(df.columns)}")
    log.info("  → %d rows, using column '%s'", len(df), col)
    df["__desc_col__"] = df[col].fillna("").astype(str).str.strip()
    return df, col


def write_results(results: pd.DataFrame, output_path):
    output_path = Path(output_path)
    log.info("Writing results to '%s' …", output_path)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        results.to_excel(writer, sheet_name="Matched Results", index=False)

        total = len(results)
        rows = []
        for m in ["Fuzzy", "TF-IDF", "Semantic", "None"]:
            cnt = (results["Method"] == m).sum()
            rows.append({
                "Method": m,
                "Count": cnt,
                "Pct": f"{round(cnt / total * 100, 1)}%" if total else "0%",
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name="Summary", index=False)

        w_rows = []
        for w in ["With Weight", "Without Weight", "N/A (no weight on either side)"]:
            cnt = (results["Weight Match"] == w).sum()
            w_rows.append({
                "Weight Match": w,
                "Count": cnt,
                "Pct": f"{round(cnt / total * 100, 1)}%" if total else "0%",
            })
        pd.DataFrame(w_rows).to_excel(writer, sheet_name="Weight Summary", index=False)

        un = results[results["Match Status"] == "Unmatched"]
        if not un.empty:
            un.to_excel(writer, sheet_name="Unmatched", index=False)

    log.info("Done. %d rows written.", len(results))


# ─────────────────────────────────────────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def run_pipeline(cfg: dict):
    t0 = time.time()
    th = cfg["thresholds"]
    layers = cfg["layers"]
    topk = cfg["top_k"]
    perf = cfg["performance"]
    weight_rule = cfg.get("weight_rule", "strict")
    sample_n = cfg.get("sample")          # ← 100 or None

    fuzzy_th = th["fuzzy"]
    tfidf_th = th["tfidf"] / 100
    sem_th   = th["semantic"] / 100

    # ── Source ──────────────────────────────────────────────────────────────
    src_cfg = cfg["source"]
    src_df, src_col = read_file(src_cfg["file"], src_cfg["column"])
    source_raw = src_df["__desc_col__"].tolist()

    log.info("Parsing %d source descriptions …", len(source_raw))
    source_parsed = []
    for raw in source_raw:
        name, pack, weight = parse_description(raw)
        source_parsed.append({
            "raw":      raw,
            "name_norm": normalize_name(name),
            "pack":     normalize_pack(pack),
            "pack_raw": pack,
            "weight_raw": weight,
            "weight_n": normalize_weight(weight),
        })
    source_names = [p["name_norm"] for p in source_parsed]

    # ── Indexes ────────────────────────────────────────────────────────────
    tfidf    = TFIDFMatcher(source_names) if layers["tfidf"] else None
    semantic = (
        SemanticMatcher(source_names, perf["semantic_model"], perf["semantic_batch"])
        if layers["semantic"] else None
    )

    # ── Targets ────────────────────────────────────────────────────────────
    all_results = []

    for tgt_cfg in cfg["targets"]:
        tgt_df, tgt_col = read_file(tgt_cfg["file"], tgt_cfg["column"])

        # ── SAMPLE LIMIT ────────────────────────────────────────────────────
        if sample_n is not None:
            tgt_df = tgt_df.head(sample_n)
            log.info("  → SAMPLE MODE: processing first %d rows only "
                     "(set 'sample': None in CONFIG for full run)", sample_n)
        # ────────────────────────────────────────────────────────────────────

        target_raw = tgt_df["__desc_col__"].tolist()
        file_name  = Path(tgt_cfg["file"]).name

        log.info("Parsing %d target descriptions …", len(target_raw))
        target_parsed = []
        for raw in target_raw:
            name, pack, weight = parse_description(raw)
            target_parsed.append({
                "raw":      raw,
                "name_norm": normalize_name(name),
                "pack":     normalize_pack(pack),
                "weight_raw": weight,
                "weight_n": normalize_weight(weight),
            })

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
            cust_raw  = tp["raw"]
            cust_name = tp["name_norm"]
            cust_w    = tp["weight_n"]
            cust_p    = tp["pack"]

            best = {"idx": None, "score": 0.0, "method": "None", "weight_ok": False}

            # Layer 2: Fuzzy
            if layers["fuzzy"] and cust_name:
                cands = fuzzy_top_k(cust_name, source_names, topk["fuzzy"], fuzzy_th)
                idx, sc, w_ok = select_best_candidate(
                    cands, cust_w, cust_p, source_parsed, weight_rule)
                if idx is not None:
                    best = {"idx": idx, "score": sc, "method": "Fuzzy", "weight_ok": w_ok}

            # Layer 3: TF-IDF
            if best["idx"] is None and tfidf is not None and cust_name:
                cands = tfidf.top_k(cust_name, topk["tfidf"], tfidf_th)
                idx, sc, w_ok = select_best_candidate(
                    cands, cust_w, cust_p, source_parsed, weight_rule)
                if idx is not None:
                    best = {"idx": idx, "score": sc, "method": "TF-IDF", "weight_ok": w_ok}

            # Layer 4: Semantic
            if best["idx"] is None and semantic is not None and cust_name:
                cands = semantic.top_k_for_embedding(
                    target_emb[i], topk["semantic"], sem_th)
                idx, sc, w_ok = select_best_candidate(
                    cands, cust_w, cust_p, source_parsed, weight_rule)
                if idx is not None:
                    best = {"idx": idx, "score": sc, "method": "Semantic", "weight_ok": w_ok}

            if best["idx"] is not None:
                sp = source_parsed[best["idx"]]
                rows.append({
                    "Source File":         file_name,
                    "Our Description":     sp["raw"],
                    "Customer Description": cust_raw,
                    "Match Status":        "Matched",
                    "Score":               best["score"],
                    "Method":              best["method"],
                    "Weight Match":        weight_match_label(sp["weight_n"], cust_w),
                })
            else:
                rows.append({
                    "Source File":         file_name,
                    "Our Description":     "",
                    "Customer Description": cust_raw,
                    "Match Status":        "Unmatched",
                    "Score":               0.0,
                    "Method":              "None",
                    "Weight Match":        (
                        "N/A (no weight on either side)"
                        if cust_w is None else "Without Weight"
                    ),
                })

        result_df = pd.DataFrame(rows)
        all_results.append(result_df)

        matched = (result_df["Match Status"] == "Matched").sum()
        log.info("  '%s' → %d/%d matched (%.1f%%)",
                 file_name, matched, len(result_df),
                 matched / len(result_df) * 100 if len(result_df) else 0)

    final = pd.concat(all_results, ignore_index=True)

    total         = len(final)
    matched_total = (final["Match Status"] == "Matched").sum()
    log.info("=" * 60)
    log.info("SUMMARY")
    log.info("  Total       : %d", total)
    log.info("  Matched     : %d (%.1f%%)",
             matched_total, matched_total / total * 100 if total else 0)
    for m in ["Fuzzy", "TF-IDF", "Semantic"]:
        log.info("    %-9s: %d", m, (final["Method"] == m).sum())
    log.info("  With Weight : %d", (final["Weight Match"] == "With Weight").sum())
    log.info("  Without Wt  : %d", (final["Weight Match"] == "Without Weight").sum())
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

    log.info("MatchIQ v2 starting …")
    log.info("  Layers      : fuzzy=%s  tfidf=%s  semantic=%s",
             cfg["layers"]["fuzzy"], cfg["layers"]["tfidf"], cfg["layers"]["semantic"])
    log.info("  Weight rule : %s", cfg.get("weight_rule", "strict"))
    log.info("  Sample limit: %s",
             str(cfg.get("sample")) + " rows" if cfg.get("sample") else "FULL FILE")
    run_pipeline(cfg)


if __name__ == "__main__":
    main()