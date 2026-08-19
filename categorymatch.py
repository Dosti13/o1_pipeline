"""semantic_category_match_v5.py  (v5 - fixes dairy false-positives + tea-bag misses)
---------------------------------------------------------------------
Builds on v4. This version fixes three real bugs found in QA:

BUGFIX 0 -- SYNTAX ERROR
---------------------------------------------------------------------
v4's FORM_WORDS["masala"] list had a stray double comma
(`... , , ...`) which is an outright SyntaxError -- the script could
not even be imported. Fixed.

BUGFIX 1 -- DAIRY PRODUCT FALSE POSITIVES (wafers/biscuits/candy)
---------------------------------------------------------------------
The Pass-3 keyword net matched "MILK" anywhere in material_name and
routed it to DAIRY PRODUCT, excluding only rows containing the full
word "CHOCOLATE". Real data abbreviates this to "CHOCO"
("...MILK CHOCO BDAY...", "...CHOCO FLAV MILK...") which does NOT
match "CHOCOLAT", so wafers, biscuits, cream-filled candy, and
chocolate novelty items ("3D ME TO YOU MILK CHOCO...") were wrongly
tagged as DAIRY PRODUCT. Also, cases like "AERO MEDIUM MILK" (a candy
tablet brand) have no chocolate wording in the name at all, but the
row's OWN category field already says "TABLETS" -- a free signal v4
ignored because the exclude check only looked at material_name.

Fix:
  - Widened the exclude pattern: CHOCO (not just CHOCOLAT), WAFER,
    BISCUIT, TABLET, CANDY/CANDIES, FILLED/FILLING, GATEAU, FUDGE,
    TOFFEE, PRALINE, CRUNCH, CARAMEL.
  - The keyword-rule check now also looks at the row's own `category`
    and `brand` text, not just material_name, so "AERO MEDIUM MILK"
    (category="TABLETS") is correctly excluded from DAIRY PRODUCT.

BUGFIX 2 -- "TEA BAG" ROWS NOT LANDING ON A TEA CATEGORY
---------------------------------------------------------------------
Pass 2b (category-name vs LULU category-name) relied purely on
embedding similarity. Noisy strings like
"HERBAL ( INFUSION / OTHERS ) TEA BAG" normalize down to
"HERBAL INFUSION OTHERS TEA BAG" -- the extra descriptor words dilute
the embedding enough to miss P2_THRESHOLD or lose the margin check,
so many tea-bag rows fell through to NULL instead of TEA BAG /
GREEN TEA / BLACK TEA BAG / SPECIALITY TEA.

Fix: added a deterministic tea-bag sub-type mapper that runs BEFORE
the embedding matcher in Pass 2b. It only fires when "TEA" appears in
the (normalized) category text, so it can't hijack unrelated
categories. Anything it doesn't confidently sub-type still falls
through to the embedding matcher as before.

Install deps:
    pip install psycopg2-binary sqlalchemy pandas sentence-transformers scikit-learn python-dotenv

Run:
    python semantic_category_match_v5.py
---------------------------------------------------------------------
"""

import os
import re
from collections import defaultdict

import numpy as np
import pandas as pd
from sqlalchemy import create_engine, text
from sentence_transformers import SentenceTransformer, util
import dotenv
dotenv.load_dotenv()
# ======================================================================
# 1. CONFIG
# ======================================================================
DB_CONFIG = {
    "host":     os.getenv("PG_HOST"),
    "port":     os.getenv("PGPORT", "5432"),
    "dbname":  os.getenv("PG_DB"),
    "user":     os.getenv("PG_HOST")    ,
    "password": os.getenv("PG_PASSWORD"),
}

SOURCE_TABLE   = "semantic.material_master_customer"
STAGING_TABLE  = "semantic.new_category_matches"
VIEW_NAME      = "semantic.vw_material_master_with_new_category"

MODEL_NAME = "all-MiniLM-L6-v2"
BATCH_SIZE = 256

# ----------------------------------------------------------------------
# MASTER CATEGORY VOCABULARY
# ----------------------------------------------------------------------
LULU_CATEGORIES_MASTER = [
    "JASMINE RICE", "ORGANIC EGGS", "MEXICAN FOODS", "BLACK LOOSE TEA",
    "CANNED BEANS", "VERMICELLI", "SLICED CHEESE", "MINCED MUTTON/LAMB",
    "CANNED FOUL BEANS", "CANNED LUNCHEON MEAT", "CANOLA OIL",
    "MIX VEGETABLE", "BROWN EGGS", "COFFEE CAPSULES", "GRAPE LEAVES",
    "ICE CREAM IMPULSE", "FROZEN WHOLE FISH", "COFFEE", "SPORTS NUTRITION",
    "CHEESE SPREADS", "BLACK TEA", "CREAM FILLED BISCUIT",
    "CANNED PINEAPPLE", "IH  OLIVES", "COOKIES", "IH CHEESE",
    "SPARKLING WATER", "BEANS", "OTHER BURGERS", "OTHER CRISPS",
    "FROZEN PIZZAS", "ZINGERS", "CHICKEN BURGERS", "FISH FINGERS & STEAK",
    "BASMATI", "KARAK & TEA MIXES", "UNKNOWN", "TEA BAG", "LABNEH",
    "CANNED TUNA", "COOKED CHICKEN", "COUSCOUS", "CANNED BAKED BEANS",
    "IH SALAD & MARINATED", "SAUSAGES PREPACKED", "CHILLED COFFEE DRINK",
    "CEREAL BARS", "POP CORN", "ETHNIC READY MEALS", "BLACK TEA BAG",
    "POTATO BAGS", "HARD CHEESE", "INSTANT COFFEE", "GREEN TEA",
    "CANNED CORNED BEEF", "FIBER BISCUITS", "OTHER EGGS",
    "OTHR.FROZ. VEGETABLE", "SOFT & GRATED CHEESE", "MINCED BEEF & VEAL",
    "VEAL & BEEF", "QUAILS", "CANND FRUIT COCKTAIL", "JAR CHEESE",
    "ROAST & GROUND COFFE", "CORN", "SHRIMPS", "RUSKS", "PRAWNS",
    "CANNED PEAS", "PORTION CHEESE", "GREEN PEAS", "ICE CREAM TAKE HOME",
    "FROZEN SAUSAGES", "BEEF BURGERS", "CAKE & DESSERT MIXES",
    "CAKES & PIES", "FRUITS", "WHITE EGGS", "COOKING SAUCE", "PASTA",
    "EGYPTIAN RICE", "SOFT CHEESE", "CANNED SAUSAGES",
    "CAND TOMATOES&PUREE", "OTHER COOKING OIL", "CHICKEN PORTIONS",
    "SP&OTHRPOTAT.PRODUCT", "SPECIALITY TEA", "OTHER CAND.VEGETABLE",
    "CAULIFLOWER", "CAKE DECORATIONS", "WHOLE CHICKENS", "NUGGETS",
    "CREAM CARAMEL", "CUSTARD", "INDIV.QUICK FROZEN", "SHARING PACKS",
    "FLOUR", "WHITE SUGAR", "CAKES", "CAND HAMMOUS&THAHINA", "DATES",
    "DIABETIC&HEALTH CAKE", "TURKEY PREPACKED", "CAKES & GATEAUX",
    "OLIVE OIL", "BEEF PREPACKED", "CAND WHL.KERNEL CORN", "VINEGAR",
    "GRATED CHEESE", "KEBABS", "BREADCRUMBS & BATTER", "MEAT BALLS",
    "CHICKEN PREPACKED", "GHEE", "SUNFLOWER OIL", "CANNED MUSHROOM",
    "OTHER FROZEN DAIRY", "DUCK", "NUTS PROCESSED", "OLIVES",
    "FRESH JUICE ASSORTED", "POPCORNS", "BLENDED OIL",
]

LULU_CATEGORIES_NEW = [
    "MASALA & MIX",
    "DAIRY PRODUCT",
]

LULU_CATEGORIES_ALL = sorted(set(LULU_CATEGORIES_MASTER) | set(LULU_CATEGORIES_NEW))

# ----------------------------------------------------------------------
# CONFIDENCE KNOBS
# ----------------------------------------------------------------------
TOP_K            = 10
P1_HARD_MIN      = 0.62
P1_AGREEMENT_MIN = 0.55
P1_MARGIN_MIN    = 0.15
P1_OVERRIDE_MIN  = 0.85

P2_THRESHOLD     = 0.60
P2_MARGIN_MIN    = 0.08

# Cross-customer borrowing: how close a same/similar product from another
# customer's catalog needs to be before we trust its category label.
P2X_NEIGHBOUR_MIN = 0.80

# ----------------------------------------------------------------------
# TEXT NORMALIZATION  (fixes plural/singular + noisy category prefixes)
# ----------------------------------------------------------------------
_PREFIX_RE = re.compile(r'^(F6_|IH\s+)', re.IGNORECASE)


def normalize_label_text(text_in: str) -> str:
    """Clean a category label before embedding: strip known noise
    prefixes, turn separators into spaces, collapse whitespace. Applied
    to BOTH source categories and LULU categories so comparisons are
    apples-to-apples."""
    if not text_in:
        return text_in
    c = _PREFIX_RE.sub('', text_in.strip())
    c = c.replace('(', ' ').replace(')', ' ')
    c = c.replace('/', ' ').replace('_', ' ').replace('&', ' and ')
    c = re.sub(r'\s+', ' ', c).strip()
    return c


def singularize(word: str) -> str:
    w = word.lower()
    if w.endswith('ies') and len(w) > 4:
        return w[:-3] + 'y'
    if w.endswith('es') and len(w) > 4 and w[-3] in 'sxzh':
        return w[:-2]
    if w.endswith('s') and not w.endswith('ss') and len(w) > 3:
        return w[:-1]
    return w


def normalized_tokens(text_in: str):
    return {singularize(t) for t in re.findall(r'[a-z]+', text_in.lower())}


# ----------------------------------------------------------------------
# FORM-WORD GUARD  (now singular/plural aware via normalized_tokens)
# ----------------------------------------------------------------------
FORM_WORDS = {
    "jam":        ["jam", "preserve", "marmalade"],
    "masala":     ["masala", "spice mix", "seasoning", "chillies", "spices"],
    "sauce":      ["sauce"],
    "ketchup":    ["ketchup"],
    "syrup":      ["syrup"],
    "paste":      ["paste"],
    "powder":     ["powder"],
    "seasoning":  ["seasoning", "masala"],
    "marinade":   ["marinade", "masala"],
    "pickle":     ["pickle", "pickled"],
    "mix":        ["mix"],
    "olives":     ["olv", "olives"],
}


def _find_form_word(name_lower: str):
    tokens = normalized_tokens(name_lower)
    for key in FORM_WORDS:
        if singularize(key) in tokens:
            return key
    return None


def _neighbours_support_form(form_key: str, neighbour_names_lower):
    synonyms = FORM_WORDS[form_key]
    for n in neighbour_names_lower:
        n_tokens = normalized_tokens(n)
        if any(singularize(s) in n_tokens or s in n for s in synonyms):
            return True
    return False


# ----------------------------------------------------------------------
# KEYWORD SAFETY NET  (Pass 3 -- explicit, narrow, word-boundary guarded)
# ----------------------------------------------------------------------
# (match_regex, exclude_regex_or_None, target_category)
# Evaluated in order; first match wins. Only runs on rows still
# unresolved after Pass 1 and Pass 2. `exclude` is now checked against
# a combined context string (material_name + own category + brand),
# not material_name alone -- see apply_keyword_rules().
CONFECTIONERY_EXCLUDE_RE = re.compile(
    r'CHOCOLAT|\bCHOCO\b|CADBURY|POWDER|COCONUT|'
    r'\bWAFER|\bBISCUIT|\bTABLET|\bCANDY|\bCANDIES|'
    r'\bFILLED\b|\bFILLING\b|\bGATEAU|\bFUDGE|\bTOFFEE|\bPRALINE|\bCRUNCH|\bCARAMEL',
    re.I,
)

KEYWORD_RULES = [
    # cake / cupcake -> CAKES, but not cake MIXES or cake DECORATIONS
    (re.compile(r'\bCUP\s?CAKE(S)?\b|\bCAKE(S)?\b', re.I),
     re.compile(r'\bMIX(ES)?\b|\bDECORAT|\bGATEAU|\bPIE(S)?\b', re.I),
     "CAKES"),
    # fresh/full-fat/skim milk, yoghurt -> DAIRY PRODUCT, but not
    # chocolate novelty items (incl. "CHOCO" abbreviation), milk
    # powder, coconut milk, or wafer/biscuit/candy/tablet products
    # that merely mention "milk" as a flavour.
    (re.compile(r'\bMILK\b|\bYOGH?URT\b|\bYOGH?OURT\b', re.I),
     CONFECTIONERY_EXCLUDE_RE,
     "DAIRY PRODUCT"),
]


def apply_keyword_rules(material_name: str, own_category: str = None, brand: str = None):
    """Runs the keyword net against material_name for the trigger match,
    but checks the exclude pattern against the FULL context (name + own
    category + brand) so signals already sitting in the row (e.g. an
    own category of "TABLETS" for a candy brand) aren't ignored."""
    if not material_name:
        return None
    context = " ".join(filter(None, [material_name, own_category or "", brand or ""]))
    for pattern, exclude, category in KEYWORD_RULES:
        if pattern.search(material_name) and not (exclude and exclude.search(context)):
            return category
    return None


# ----------------------------------------------------------------------
# TEA BAG SUB-TYPE MAPPER  (deterministic, runs before Pass 2b embedding)
# ----------------------------------------------------------------------
# Noisy category strings that clearly reference tea/tea bags often lose
# too much embedding similarity against clean LULU labels once they
# carry extra descriptor words ("HERBAL ( INFUSION / OTHERS ) TEA BAG").
# Resolve these deterministically instead of leaving it to embeddings.
TEA_BAG_SUBTYPE_RULES = [
    (re.compile(r'\bGREEN\b', re.I), "GREEN TEA"),
    (re.compile(r'\bBLACK\b', re.I), "BLACK TEA BAG"),
    (re.compile(r'HERBAL|INFUSION|CHAMOMILE|CAMOMILE|MINT|ROOIBOS|LEMON|GINGER|'
                r'SLIM|DIGEST|IMMUNE|HIBISCUS|SAGE', re.I),
     "SPECIALITY TEA"),
]


def map_tea_bag_category(cat_text: str):
    """Returns a LULU tea category for a source category string that
    clearly references tea / tea bags, or None if it doesn't apply
    (letting the embedding matcher in Pass 2b handle it as before)."""
    if not cat_text:
        return None
    norm = normalize_label_text(cat_text).upper()
    if "TEA" not in norm:
        return None
    for pattern, target in TEA_BAG_SUBTYPE_RULES:
        if pattern.search(norm):
            return target
    if "TEA BAG" in norm:
        return "TEA BAG"
    if "TEA" in norm:
        # Loose leaf / unspecified tea references without "BAG" and
        # without a recognizable sub-type -- let embeddings decide
        # rather than guessing.
        return None
    return None


# ======================================================================
# 2. DB CONNECTION
# ======================================================================
def get_engine():
    conn_str = (
        f"postgresql+psycopg2://{DB_CONFIG['user']}:{DB_CONFIG['password']}"
        f"@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['dbname']}"
    )
    return create_engine(conn_str)


def load_data(engine) -> pd.DataFrame:
    query = f"""
        SELECT material_code, material_name, brand, category, source_customer
        FROM {SOURCE_TABLE}
        WHERE material_name IS NOT NULL
    """
    df = pd.read_sql(text(query), engine)
    df["source_customer"] = df["source_customer"].str.strip().str.upper()
    df["material_name"] = df["material_name"].str.strip()
    df["category"] = df["category"].str.strip()
    return df


# ======================================================================
# 3. SEMANTIC MATCHING
# ======================================================================
def build_lulu_reference(df: pd.DataFrame):
    lulu = df[df["source_customer"] == "LULU"].dropna(subset=["material_name", "category"])
    lulu = lulu.drop_duplicates(subset=["material_name"])

    before = len(lulu)
    lulu = lulu[lulu["category"].isin(LULU_CATEGORIES_MASTER)]
    dropped = before - len(lulu)
    if dropped:
        print(f"Note: dropped {dropped} LULU reference rows with categories "
              f"not in the master vocabulary (likely typos/one-offs).")

    return lulu.reset_index(drop=True), LULU_CATEGORIES_ALL


def encode_texts(model, texts):
    return model.encode(
        texts,
        batch_size=BATCH_SIZE,
        convert_to_tensor=True,
        show_progress_bar=True,
        normalize_embeddings=True,
    )


def semantic_match_voted(model, query_texts, ref_texts, ref_labels,
                         top_k, hard_min, agreement_min, margin_min, override_min):
    query_emb = encode_texts(model, query_texts)
    ref_emb = encode_texts(model, ref_texts)
    hits = util.semantic_search(query_emb, ref_emb, top_k=top_k)

    labels, top1s, confs = [], [], []
    for query_name, hit in zip(query_texts, hits):
        query_lower = query_name.lower()
        form_key = _find_form_word(query_lower)

        if not hit:
            labels.append(None); top1s.append(0.0); confs.append(0.0)
            continue

        top1 = float(hit[0]["score"])

        votes = defaultdict(float)
        names_by_label = defaultdict(list)
        for n in hit:
            lbl = ref_labels[n["corpus_id"]]
            votes[lbl] += float(n["score"])
            names_by_label[lbl].append(ref_texts[n["corpus_id"]].lower())

        ranked = sorted(votes.items(), key=lambda kv: kv[1], reverse=True)
        win_label, win_vote = ranked[0]
        runner_vote = ranked[1][1] if len(ranked) > 1 else 0.0
        total = sum(votes.values()) or 1.0

        confidence = win_vote / total
        margin = (win_vote - runner_vote) / total

        passes_vote_gate = (top1 >= hard_min
                            and confidence >= agreement_min
                            and margin >= margin_min)
        passes_override = top1 >= override_min

        ok = passes_vote_gate or passes_override

        if ok and form_key is not None:
            if not _neighbours_support_form(form_key, names_by_label[win_label]):
                ok = False

        labels.append(win_label if ok else None)
        top1s.append(top1)
        confs.append(confidence)

    return labels, top1s, confs


def semantic_match_margin(model, query_texts, ref_texts, ref_labels,
                          threshold, margin_min):
    """Category-name <-> LULU category-name matcher. Both sides are
    normalized (F6_ prefixes stripped, punctuation cleaned) before
    embedding so noisy source labels compare fairly against clean LULU
    labels."""
    norm_query = [normalize_label_text(t) for t in query_texts]
    norm_ref = [normalize_label_text(t) for t in ref_texts]

    query_emb = encode_texts(model, norm_query)
    ref_emb = encode_texts(model, norm_ref)
    hits = util.semantic_search(query_emb, ref_emb, top_k=2)

    labels, scores = [], []
    for hit in hits:
        if not hit:
            labels.append(None); scores.append(0.0)
            continue
        top1 = float(hit[0]["score"])
        top2 = float(hit[1]["score"]) if len(hit) > 1 else 0.0
        ok = top1 >= threshold and (top1 - top2) >= margin_min
        labels.append(ref_labels[hit[0]["corpus_id"]] if ok else None)
        scores.append(top1)
    return labels, scores


def semantic_match_margin_with_tea_override(model, query_texts, ref_texts, ref_labels,
                                            threshold, margin_min):
    """Same contract as semantic_match_margin, but first tries the
    deterministic tea-bag sub-type mapper on each query text. Only
    texts the mapper declines to handle go through the embedding
    matcher, avoiding the dilution problem noisy tea-bag category
    strings ran into in v4."""
    labels = [None] * len(query_texts)
    scores = [None] * len(query_texts)

    embed_idx, embed_texts = [], []
    for i, cat_text in enumerate(query_texts):
        direct = map_tea_bag_category(cat_text)
        if direct is not None:
            labels[i] = direct
            scores[i] = 1.0  # deterministic match, treat as fully confident
        else:
            embed_idx.append(i)
            embed_texts.append(cat_text)

    if embed_texts:
        sub_labels, sub_scores = semantic_match_margin(
            model, embed_texts, ref_texts, ref_labels,
            threshold=threshold, margin_min=margin_min,
        )
        for idx, lbl, sc in zip(embed_idx, sub_labels, sub_scores):
            labels[idx] = lbl
            scores[idx] = sc

    return labels, scores


def borrow_cross_customer_categories(model, unresolved_names, all_rows_df):
    """
    Pass 2, step A: for names Pass 1 couldn't resolve, look at every
    non-LULU customer's catalog (not just the row's own customer) for the
    nearest material_name that DOES have a category filled in, and borrow
    that label. This covers rows whose own category is blank (common in
    this data) by leaning on how a similar/identical product was labeled
    elsewhere.

    Returns a dict: material_name -> borrowed_category_text (or None).
    """
    donor_pool = all_rows_df[
        (all_rows_df["source_customer"] != "LULU")
        & all_rows_df["category"].notna()
        & (all_rows_df["category"].str.len() > 0)
    ].drop_duplicates(subset=["material_name"])

    if donor_pool.empty or not unresolved_names:
        return {}

    donor_names = donor_pool["material_name"].tolist()
    donor_cats = donor_pool["category"].tolist()

    query_emb = encode_texts(model, unresolved_names)
    donor_emb = encode_texts(model, donor_names)
    hits = util.semantic_search(query_emb, donor_emb, top_k=1)

    borrowed = {}
    for name, hit in zip(unresolved_names, hits):
        if not hit:
            continue
        score = float(hit[0]["score"])
        if score >= P2X_NEIGHBOUR_MIN:
            borrowed[name] = donor_cats[hit[0]["corpus_id"]]
    return borrowed


def run_matching(df: pd.DataFrame) -> pd.DataFrame:
    print("Loading embedding model:", MODEL_NAME)
    model = SentenceTransformer(MODEL_NAME)

    lulu_ref, lulu_categories = build_lulu_reference(df)
    print(f"LULU reference materials: {len(lulu_ref)}  |  unique categories: {len(lulu_categories)}")

    non_lulu = df[df["source_customer"] != "LULU"].copy()
    print(f"Non-LULU rows to match: {len(non_lulu)}")

    unique_names = non_lulu["material_name"].dropna().unique().tolist()
    print(f"Unique non-LULU material_names: {len(unique_names)}")

    # ---- Pass 1: material_name vs LULU material_name (voted) ----
    print("\n[Pass 1] Voted semantic match + form-word guard + high-conf override")
    labels_1, top1_1, conf_1 = semantic_match_voted(
        model,
        unique_names,
        lulu_ref["material_name"].tolist(),
        lulu_ref["category"].tolist(),
        top_k=TOP_K,
        hard_min=P1_HARD_MIN,
        agreement_min=P1_AGREEMENT_MIN,
        margin_min=P1_MARGIN_MIN,
        override_min=P1_OVERRIDE_MIN,
    )
    pass1_map = pd.DataFrame({
        "material_name": unique_names,
        "match_category_p1": labels_1,
        "match_score_p1": top1_1,
        "match_conf_p1": conf_1,
    })

    unresolved_names = pass1_map.loc[pass1_map["match_category_p1"].isna(), "material_name"].tolist()

    # ---- Pass 2, step A: borrow category from a similar product at ANY
    #      other customer (fixes rows with a blank own-category field) ----
    print(f"\n[Pass 2a] Cross-customer category borrowing for "
          f"{len(unresolved_names)} unresolved names")
    borrowed = borrow_cross_customer_categories(model, unresolved_names, df)
    print(f"  Borrowed a category label for {len(borrowed)} names")

    # ---- Pass 2, step B: own category (fallback when no good donor) ----
    own_cat_lookup = (
        non_lulu[non_lulu["material_name"].isin(unresolved_names) & non_lulu["category"].notna()]
        .drop_duplicates(subset=["material_name"])
        .set_index("material_name")["category"]
        .to_dict()
    )

    candidate_label_texts = {}
    for name in unresolved_names:
        if name in borrowed:
            candidate_label_texts[name] = borrowed[name]
        elif name in own_cat_lookup:
            candidate_label_texts[name] = own_cat_lookup[name]

    if candidate_label_texts:
        names_for_p2 = list(candidate_label_texts.keys())
        cat_texts_for_p2 = list(candidate_label_texts.values())
        print(f"\n[Pass 2b] Category-name match against LULU master list "
              f"({len(names_for_p2)} candidates) -- tea-bag deterministic "
              f"override first, embeddings for the rest")
        labels_2, scores_2 = semantic_match_margin_with_tea_override(
            model, cat_texts_for_p2, lulu_categories, lulu_categories,
            threshold=P2_THRESHOLD, margin_min=P2_MARGIN_MIN,
        )
        pass2_map = pd.DataFrame({
            "material_name": names_for_p2,
            "match_category_p2": labels_2,
            "match_score_p2": scores_2,
        })
    else:
        pass2_map = pd.DataFrame(columns=["material_name", "match_category_p2", "match_score_p2"])

    result = non_lulu.merge(pass1_map, on="material_name", how="left")
    result = result.merge(pass2_map, on="material_name", how="left")

    # ---- Pass 3: keyword safety net for names still unresolved ----
    # Now uses full row context (material_name + own category + brand)
    # so signals like an own category of "TABLETS" correctly exclude
    # candy items from the DAIRY PRODUCT rule.
    still_unresolved_mask = result["match_category_p1"].isna() & result["match_category_p2"].isna()
    result["match_category_p3"] = None
    result.loc[still_unresolved_mask, "match_category_p3"] = result.loc[still_unresolved_mask].apply(
        lambda r: apply_keyword_rules(r["material_name"], r.get("category"), r.get("brand")),
        axis=1,
    )
    n_p3 = result["match_category_p3"].notna().sum()
    print(f"\n[Pass 3] Keyword safety net matched {n_p3} rows")

    def pick(row):
        if pd.notna(row.get("match_category_p1")):
            return (row["match_category_p1"], row["match_score_p1"],
                    row.get("match_conf_p1"), "MATERIAL_NAME")
        if pd.notna(row.get("match_category_p2")):
            return (row["match_category_p2"], row["match_score_p2"],
                    None, "CATEGORY_NAME")
        if pd.notna(row.get("match_category_p3")):
            return (row["match_category_p3"], None, None, "KEYWORD_RULE")
        best = max(row.get("match_score_p1") or 0, row.get("match_score_p2") or 0)
        return (None, best, row.get("match_conf_p1"), None)

    picked = result.apply(pick, axis=1, result_type="expand")
    picked.columns = ["new_category", "similarity_score", "agreement", "match_source"]
    result = pd.concat([result[["material_code"]], picked], axis=1)

    return result


# ======================================================================
# 4. WRITE RESULTS BACK TO POSTGRES + CREATE VIEW
# ======================================================================
def write_results(engine, result: pd.DataFrame):
    print(f"Dropping dependent view {VIEW_NAME} (if it exists) ...")
    with engine.begin() as conn:
        conn.execute(text(f"DROP VIEW IF EXISTS {VIEW_NAME}"))

    print(f"\nWriting {len(result)} rows to staging table {STAGING_TABLE} ...")
    schema, table = STAGING_TABLE.split(".")
    result.to_sql(
        table, engine, schema=schema, if_exists="replace",
        index=False, method="multi", chunksize=1000,
    )

    print(f"Creating view {VIEW_NAME} ...")
    view_sql = f"""
        CREATE OR REPLACE VIEW {VIEW_NAME} AS
        SELECT
            m.material_code, m.material_name, m.brand, m.category,
            m.source_customer, s.new_category, s.similarity_score,
            s.agreement, s.match_source
        FROM {SOURCE_TABLE} m
        LEFT JOIN {STAGING_TABLE} s ON s.material_code = m.material_code
        ORDER BY m.source_customer, m.material_name;
    """
    with engine.begin() as conn:
        conn.execute(text(view_sql))
    print("Done. View created:", VIEW_NAME)


# ======================================================================
# 5. MAIN
# ======================================================================
def main():
    engine = get_engine()
    df = load_data(engine)
    result = run_matching(df)

    matched = result["new_category"].notna().sum()
    total = len(result)
    print(f"\nMatched (confident): {matched} / {total}  ({matched/total:.1%})")
    print(f"NULL (low confidence): {total - matched} / {total}  ({(total-matched)/total:.1%})")
    print("\nBy match source:")
    print(result["match_source"].value_counts(dropna=False))

    write_results(engine, result)


if __name__ == "__main__":
    main()