"""
LULU-AS-MASTER SKU MATCHING PIPELINE — v1.0
=============================================
Adapted from QNIE SKU Matching Pipeline v7.1.

Difference from v7.1:
  - Master/reference set = LULU rows from semantic.material_master_customer
    (instead of semantic.dim_material_master)
  - Customers matched against Lulu = TALABAT, C4, GRAND_MALL, AL_MEERA
    (spar excluded - insufficient data quality; lulu excluded - it IS the master)
  - Single source table for everything: semantic.material_master_customer,
    filtered by source_customer. No separate staging tables, no billing bridge
    (the old bridge view is keyed to the QNIE master's material_code, which has
    no meaning here, so it's dropped entirely).
  - Output: CSV only (no DB write).

Run this where you already run etl_job_qnie.py / fuzzy4.py — it needs psycopg2
and a network path to 10.250.160.72.
"""

import os
import re
import sys
import logging
import datetime as dt
import pickle
import numpy as np
from pathlib import Path


def _ensure(pkg, import_as=None):
    name = import_as or pkg
    try:
        __import__(name)
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg,
                               "--quiet", "--break-system-packages"])


_ensure("sentence-transformers", "sentence_transformers")
_ensure("faiss-cpu", "faiss")
_ensure("rapidfuzz")
_ensure("psycopg2-binary", "psycopg2")
_ensure("pandas")

from sentence_transformers import SentenceTransformer
import faiss
from rapidfuzz import fuzz

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# =====================================================================
# CONFIG — adjust CUSTOMER_KEYS after the first run shows you the real
# distinct source_customer values (script prints them up front).
# =====================================================================
CONFIG = {
    "pg_host":     os.environ.get("PG_HOST",     "10.250.160.72"),
    "pg_port":     int(os.environ.get("PG_PORT", "5432")),
    "pg_database": os.environ.get("PG_DB",       "qnie_new_test"),
    "pg_user":     os.environ.get("PG_USER",     "postgres"),
    "pg_password": os.environ.get("PG_PASS",     "postgres"),

    "source_table": "semantic.material_master_customer",
    "master_key":    "LULU",
    "customer_keys": ["TALABAT", "C4", "GRAND_MALL", "AL_MEERA"],  # <-- verify against printed DISTINCT list

    "csv_output_path": os.environ.get("CSV_OUTPUT", "lulu_match_results.csv"),

    "model_name":           "all-MiniLM-L6-v2",
    "embedding_batch_size": 256,
    "faiss_top_k":          10,
    "ai_threshold_matched":    0.80,
    "brand_fuzzy_threshold":   85,
    "classify_min_score":      80.0,

    "cache_dir": "./cache_lulu_master",
}

# =====================================================================
# Normalization / brand logic — copied as-is from v7.1 so scores stay
# comparable / consistent with the existing pipeline.
# =====================================================================
KNOWN_BRANDS: set = {
    "almarai", "nada", "baladna", "nadec", "al rawabi", "puck",
    "pinar", "sadia", "danone", "nestle", "president", "arla",
    "perla", "kdd", "ola", "anchor", "lactel", "lurpak",
    "aklena", "sutas", "lactonia", "galbani", "biomass", "smeds", "dolce",
    "lipton", "nescafe", "pepsi", "monster", "rani", "tang",
    "vimto", "mirinda", "7up", "sprite", "ahmad", "deemah",
    "rubicon", "activia", "voss", "rayyan", "jouf", "hayat", "uludag",
    "segafredo", "levista", "coca cola", "red bull",
    "iffco", "wesson", "afia", "alfa", "luna", "coroli", "aseel",
    "noor", "alard", "mazola", "colavita", "olitalia", "safi",
    "sunny", "yara", "arzco", "mima", "dalda", "nawar", "salam",
    "mumtaz", "laura", "sabah", "safya", "fortune",
    "asila", "areeq", "al theka", "camolino", "judi", "torrent",
    "semiramis", "queen", "rkg",
    "americana", "galaxy", "pringles", "lays", "doritos",
    "tiffany", "tiff", "dandy", "hilal", "nabil", "delicio",
    "bunalun", "regal", "tgif", "bauducco", "oreo", "cadbury",
    "parliament", "tonys", "hectares", "quanta", "shehrazade",
    "rc", "ld", "toblerone", "fantasia", "bc", "moms",
    "old denmark", "my bizcuit", "gpr", "lohilo", "mr organic", "sara",
    "seara", "farmland", "mccain", "zwan", "doux",
    "siblou", "rahma", "perdix", "bobo", "nowaco", "namet",
    "dmh", "felza", "qualiko", "brz", "rastelli", "duchesse",
    "foodys", "beyond", "caillor", "fisher", "pena",
    "al shamal", "mazzraty", "zowadeh", "elmaha", "niers",
    "sirella", "oceana", "c best", "gourmet",
    "nest", "natura", "zlatno",
    "ukraine", "gcfc", "athba", "a saffa", "al diar",
    "panzani", "lotus", "nurjahan", "punjab", "blossom",
    "dana", "sunwhite", "adriana", "kashkaval",
    "al mumtaz", "al naham", "devaaya", "al doha",
    "crops",
    "khazan", "sunbulah", "zain", "amc", "mb", "bf",
    "ufc", "greens", "allana", "promex", "khaburah", "sriracha",
    "maxims", "al tayyab",
    "igloo", "lutosa", "danette", "khoury",
    "cashmere", "akwa",
    "saba", "rawa", "valencia", "hala", "ghalia", "sunripe", "frooti", "ghadeer",
    "ame", "cg", "gg", "lamesa", "nv", "rs", "oep",
    "orima", "amer", "awafi", "rona", "pride",
    "wadi", "rana", "hilli", "albadia", "chtoura",
    "sunblast", "ottogi", "kaanlar", "cadburys", "turkish",
    "qatari", "mara", "tilda", "mannai", "rosary",
    "bradma", "nadine", "kwality", "agrilife", "kohinoor",
    "granoro", "datschaub", "wafi", "santa",
    "highlands", "durra", "bordon", "sparoni",
    "banetti", "baytouti", "emborg", "campagna", "watties",
    "nobar", "plein", "soleil", "bonduelle", "burcu",
    "nat", "talia", "chi", "alwadi", "hershey", "supreme",
    "acetum", "qnie", "uht",
    "lulu", "carrefour", "spar", "crf", "mcrf",
    "al ain", "al rawabi", "al islami", "al kabeer", "al wazir",
    "al jawhra", "al meera", "al marai", "al zain", "al baker",
    "al gazel", "al jazira", "al wadi", "abu bint",
    "al dana", "al noor", "al osra", "al nutrica", "al shafi",
    "al yaqeen", "al baraka",
    "happy cow", "blue marine", "green giant", "sara lee",
    "california garden", "foster clark", "punjab garden",
    "royal umbrella", "royal classic", "royal chicken", "royal tender",
    "zain farm", "rc luxury", "nature valley", "betty crocker",
    "del monte", "old town", "india gate", "indus valley",
    "chtoura fields", "cashmere saffron", "nawar sunflower",
    "salam sunflower", "dalda vegetable", "aqua gulf",
    "golden irish", "royal castle", "royal garden",
    "al rayan", "al nahla", "al douha", "al kasih",
    "al bayrouty", "al badia", "al saffa", "al naseem",
    "al alali", "al forno", "al marrai", "al waha",
    "london dairy", "granja moro", "rio mare", "future farm",
    "wadi food", "super moist", "vita coco", "cocofly",
    "happy gardens", "santa maria", "coca cola", "red bull", "kit kat",
    "mr organic", "old denmark", "my bizcuit", "c best",
    "al shamal", "al mumtaz", "al naham", "al doha",
    "al theka", "al diar", "al tayyab", "a saffa",
}

_BRAND_ABBREV: dict = {
    "ame": "americana", "amc": "americana", "amcana": "americana",
    "amcn": "americana", "ame.": "americana", "amer": "americana",
    "cg": "california garden", "cal garden": "california garden", "cal": "california garden",
    "gg": "green giant", "g5": "green giant",
    "bc": "betty crocker",
    "ld": "london dairy", "londondairy": "london dairy", "l/d": "london dairy",
    "oep": "old el paso",
    "kdc": "kdd", "kd": "kdd",
    "banetti": "panzani", "sparoni": "panzani", "panz": "panzani",
    "siella": "sirella",
    "zawan": "zwan", "zwdh": "zwan",
    "bb": "butterball",
    "hershey's": "hershey",
    "cadburys": "cadbury",
    "indian gate": "india gate",
    "prsdnt": "president",
    "nurjuhan": "nurjahan",
    "frmland": "farmland",
}

_STOPWORDS = frozenset({
    'the','a','an','and','or','of','in','for','with','by','from','to','at','on',
    'is','are','new','original','premium','special','classic','natural','fresh',
    'pure','extra','super','ultra','mini','big','large','small','best','real',
    'true','fine','frozen','chilled','hot','spicy','crispy','crunchy','regular',
    'light','full','fat','low','rich','deluxe','assorted','asstd','asst',
    'prm','prem','std','alu','tetra','cpd','chopped','whole','sliced','shredded',
    'diced','minced','bottle','jar','tub','pouch','family','traditional','artisan',
    'sugar','lean','skimmed','piece','stem','chunk','strip','sachet','squeeze',
    'spray','drinking','mineral','spring','purified','organic','bio','homestyle',
    'kids','junior','baby','adult','promo','offer','deal','halal','kosher',
    'vegan','vegetarian','gold','silver','platinum','bronze','smoked','roasted',
    'grilled','baked','boiled','cooked','marinated','seasoned','unseasoned','breaded',
})

_UNIT_PATTERNS = [
    (re.compile(r'\b(ltr|lt|litre|liter|liters|litres)\b'),  'l'),
    (re.compile(r'\b(grm|gms|gram|grams|gr|gm)\b'),           'g'),
    (re.compile(r'\b(mltr|milliliter|millilitre)\b'),          'ml'),
    (re.compile(r'\b(kilogram|kilograms)\b'),                  'kg'),
    (re.compile(r'\b(ounce|ounces)\b'),                        'oz'),
    (re.compile(r'\b(piece|pieces)\b'),                        'pcs'),
    (re.compile(r'\b(packet|packets|pkt)\b'),                  'pack'),
]

_ABBREV_PATTERNS = [
    (re.compile(r'\bamcana\b'),          'americana'),
    (re.compile(r'\b(ame|amc|amcn)\b'), 'americana'),
    (re.compile(r'\bcal garden\b'),      'california garden'),
    (re.compile(r'\b(cg)\b'),            'california garden'),
    (re.compile(r'\b(oep)\b'),           'old el paso'),
    (re.compile(r'\b(gg)\b'),            'green giant'),
    (re.compile(r'\b(bc)\b'),            'betty crocker'),
    (re.compile(r'\b(ld)\b'),            'london dairy'),
    (re.compile(r'\blondondairy\b'),     'london dairy'),
    (re.compile(r'\b(kdc|kd)\b'),        'kdd'),
    (re.compile(r'\b(ckn|chk|chkn)\b'), 'chicken'),
    (re.compile(r'\b(fz|frz)\b'),        'frozen'),
    (re.compile(r'\b(veg|veget)\b'),     'vegetable'),
    (re.compile(r'\b(choc|choco)\b'),    'chocolate'),
    (re.compile(r'\b(str|strw|straw|strbry|strwbry)\b'), 'strawberry'),
    (re.compile(r'\b(van|vnla|vanla)\b'),'vanilla'),
    (re.compile(r'\b(msh|mshm)\b'),      'mushroom'),
    (re.compile(r'\bzawan\b'),           'zwan'),
    (re.compile(r'\bsiella\b'),          'sirella'),
    (re.compile(r'\bindian gate\b'),     'india gate'),
    (re.compile(r'\bnurjuhan\b'),        'nurjahan'),
    (re.compile(r'\btawouk\b'),          'taouk'),
    (re.compile(r'\b(ornge?|orng)\b'),   'orange'),
    (re.compile(r'\b(moz|mozarella|mozzrla)\b'), 'mozzarella'),
    (re.compile(r'\b(jce|juc|nectar)\b'),'juice'),
    (re.compile(r'\bbanetti\b'),         'panzani'),
    (re.compile(r'\bsparoni\b'),         'panzani'),
    (re.compile(r'\b(rusk|bisc)\b'),     'biscuit'),
    (re.compile(r'\bnacho\b'),           'tortilla'),
    (re.compile(r'\b(gr|gm)\b'),         'g'),
    (re.compile(r'\bmedamma[s]?\b'),      'foul'),
    (re.compile(r'\bcalrose\b'),         'jasmine'),
    (re.compile(r'\b(sella|rozana)\b'),  'basmati'),
    (re.compile(r'\bprsdnt\b'),          'president'),
    (re.compile(r'\bcadburys\b'),        'cadbury'),
    (re.compile(r'\bsaralee\b'),         'sara lee'),
    (re.compile(r'\b(ec|qnie|new)\b'),   ''),
]

_PACK_RE   = re.compile(r'\b(carton|ctn|case|pack|pcs?|units?|pieces?|promo|shrink|bundle|offer|sp|foc|eoe|eof|po)\b')
_MULTI_RE  = re.compile(r'\b\d+[xX]\d+[xX]\d+[xX]?(\d+\.?\d*(?:ml|l|kg|g|oz|lb)?)\b')
_MULTI2_RE = re.compile(r'\b\d+[xX]\d+[xX](\d+\.?\d*(?:ml|l|kg|g|oz|lb)?)\b')
_SIZE_RE   = re.compile(r'(\d+\.?\d*)\s*(ml|l|kg|g|oz|lb)\b')
_XQTY_RE   = re.compile(r'\s*[xX]\s*\d+\b(?!\s*(?:ml|l|kg|g|oz|lb))')
_PCT_RE    = re.compile(r'\d+%\s*(?:off|extra|free|discount)?\b')
_FREE_RE   = re.compile(r'\(\d+\s*free\)|\b\d+\s*free\b')
_ALNUM_RE  = re.compile(r'[^\w\s]')
_SPACE_RE  = re.compile(r'\s+')


def normalize_text(text: str) -> str:
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
    for pat, repl in _ABBREV_PATTERNS:
        t = pat.sub(repl, t)
    tokens = [tok for tok in t.split() if tok not in _STOPWORDS]
    return _SPACE_RE.sub(' ', ' '.join(tokens)).strip()


def _resolve_brand_token(token: str) -> str:
    return _BRAND_ABBREV.get(token, token)


def extract_brand(norm_text: str, known_brands: set) -> tuple:
    words = norm_text.split()
    if not words:
        return None, norm_text
    if len(words) >= 3:
        three = " ".join(words[:3])
        three_r = _resolve_brand_token(three)
        if three_r in known_brands or three in known_brands:
            brand = three_r if three_r in known_brands else three
            return brand, " ".join(words[3:]).strip()
    if len(words) >= 2:
        two = words[0] + " " + words[1]
        two_r = _resolve_brand_token(two)
        if two_r in known_brands or two in known_brands:
            brand = two_r if two_r in known_brands else two
            return brand, " ".join(words[2:]).strip()
    one = words[0]
    one_r = _resolve_brand_token(one)
    if one_r in known_brands or one in known_brands:
        brand = one_r if one_r in known_brands else one
        return brand, " ".join(words[1:]).strip()
    return None, norm_text


def brands_match(brand_a, brand_b, fuzzy_threshold: int = 85) -> bool:
    if brand_a is None or brand_b is None:
        return False
    if brand_a == brand_b:
        return True
    a_r = _BRAND_ABBREV.get(brand_a, brand_a)
    b_r = _BRAND_ABBREV.get(brand_b, brand_b)
    if a_r == b_r:
        return True
    return fuzz.ratio(brand_a.replace(" ", ""), brand_b.replace(" ", "")) >= fuzzy_threshold


def auto_extend_brands(norm_texts: list, base_brands: set, min_freq: int = 5) -> set:
    from collections import Counter
    generic_skip = {
        'frozen','chilled','fresh','whole','mixed','assorted','white','brown',
        'black','red','green','yellow','small','medium','large','mini','jumbo',
        'beef','chicken','lamb','fish','prawn','turkey','milk','cream','butter',
        'cheese','yogurt','juice','rice','pasta','bread','flour','oil','water',
        'sugar','salt','pepper','sauce','paste','eggs','egg','dates','honey','jam',
        'cooked','raw','plain','sweet','salty','uht','organic','halal','natural',
        'quality','premium','classic','traditional','original','special','deluxe',
    }
    one_c, two_c = Counter(), Counter()
    for t in norm_texts:
        words = t.split()
        if not words:
            continue
        if words[0] not in generic_skip:
            one_c[words[0]] += 1
        if len(words) >= 2 and words[0] not in generic_skip and words[1] not in generic_skip:
            two_c[words[0] + " " + words[1]] += 1
    new_brands = set(base_brands)
    for w, cnt in one_c.items():
        if cnt >= min_freq and len(w) >= 2 and w not in new_brands:
            new_brands.add(w)
    for bg, cnt in two_c.items():
        if cnt >= min_freq and bg not in new_brands:
            new_brands.add(bg)
    added = len(new_brands) - len(base_brands)
    log.info(f"  Dynamic brands: +{added} new (total={len(new_brands)})")
    return new_brands


_NUM_PAT = re.compile(r'^\d')


def word_overlap_ratio(a: str, b: str) -> float:
    sa = {t for t in a.split() if not _NUM_PAT.match(t)}
    sb = {t for t in b.split() if not _NUM_PAT.match(t)}
    if not sa or not sb:
        return 0.0
    return len(sa & sb) / len(sa | sb)


def length_ratio(a: str, b: str) -> float:
    la, lb = len(a.split()), len(b.split())
    if max(la, lb) == 0:
        return 0.0
    return min(la, lb) / max(la, lb)


def rapidfuzz_score(q_norm: str, cand_norm: str) -> float:
    if word_overlap_ratio(q_norm, cand_norm) < 0.30:
        return 0.0
    s_sort = fuzz.token_sort_ratio(q_norm, cand_norm) / 100.0
    s_part = fuzz.partial_ratio(q_norm, cand_norm) / 100.0
    lr = length_ratio(q_norm, cand_norm)
    if lr < 0.5:
        s_part *= lr
    return 0.70 * s_sort + 0.30 * s_part


def brand_strict_score(q_norm, cand_norm, raw_ai_score, known_brands, fuzzy_threshold=85):
    q_brand,    q_product    = extract_brand(q_norm, known_brands)
    cand_brand, cand_product = extract_brand(cand_norm, known_brands)
    if q_brand and cand_brand:
        if brands_match(q_brand, cand_brand, fuzzy_threshold):
            overlap = word_overlap_ratio(q_product or q_norm, cand_product or cand_norm)
            if overlap < 0.30:
                return raw_ai_score * 0.72, "brand_match_low_product"
            elif overlap < 0.55:
                return raw_ai_score * 0.88, "brand_match_mid_product"
            return raw_ai_score, "brand_match"
        else:
            return 0.0, "brand_mismatch_rejected"
    if (q_brand and not cand_brand) or (not q_brand and cand_brand):
        return 0.0, "brand_one_sided_rejected"
    q_words    = q_norm.split()
    cand_words = cand_norm.split()
    if not q_words or not cand_words:
        return raw_ai_score * 0.70, "no_brand_both_empty"
    first_word_score = fuzz.ratio(q_words[0].replace("-",""), cand_words[0].replace("-",""))
    if first_word_score >= fuzzy_threshold:
        overlap = word_overlap_ratio(q_norm, cand_norm)
        if overlap < 0.30:
            return raw_ai_score * 0.70, "no_brand_same_first_low_overlap"
        elif overlap < 0.55:
            return raw_ai_score * 0.88, "no_brand_same_first_mid_overlap"
        return raw_ai_score, "no_brand_same_first"
    else:
        return 0.0, "no_brand_diff_first_rejected"


# =====================================================================
# DB access — single unified table
# =====================================================================
def get_conn():
    import psycopg2
    return psycopg2.connect(
        host=CONFIG["pg_host"], port=CONFIG["pg_port"],
        dbname=CONFIG["pg_database"],
        user=CONFIG["pg_user"], password=CONFIG["pg_password"],
        keepalives=1, keepalives_idle=30,
        keepalives_interval=10, keepalives_count=10,
        connect_timeout=30,
        options="-c statement_timeout=0 -c idle_in_transaction_session_timeout=0",
    )


def print_distinct_source_customers(conn):
    tbl = CONFIG["source_table"]
    with conn.cursor() as cur:
        cur.execute(f"SELECT DISTINCT source_customer, COUNT(*) FROM {tbl} GROUP BY source_customer ORDER BY 1")
        rows = cur.fetchall()
    log.info("  Distinct source_customer values in table:")
    for sc, cnt in rows:
        log.info(f"    {sc!r:<20} {cnt:>8,} rows")
    return {sc for sc, _ in rows}


def load_master(conn) -> tuple:
    tbl = CONFIG["source_table"]
    key = CONFIG["master_key"]
    log.info(f"  Loading MASTER (source_customer={key})...")
    with conn.cursor() as cur:
        cur.execute(f"""
            SELECT material_code::TEXT, material_name, primary_barcode::TEXT
            FROM {tbl}
            WHERE UPPER(source_customer) = UPPER(%s)
              AND material_name IS NOT NULL AND TRIM(material_name) <> ''
        """, (key,))
        rows = cur.fetchall()
    ids   = [r[0] for r in rows]
    raws  = [r[1].strip() for r in rows]
    norms = [normalize_text(r[1]) for r in rows]
    bc_lkp = {}
    for mid, bc in zip(ids, [r[2] for r in rows]):
        bc = str(bc).strip() if bc else None
        if bc and bc not in ('0', 'None', ''):
            bc_lkp[bc] = mid
    log.info(f"  Master ({key}): {len(ids):,} records | barcodes: {len(bc_lkp):,}")
    return ids, raws, norms, bc_lkp


def load_customers(conn) -> dict:
    tbl = CONFIG["source_table"]
    keys = CONFIG["customer_keys"]
    all_desc = {}
    for key in keys:
        log.info(f"  Loading customer (source_customer={key})...")
        with conn.cursor() as cur:
            cur.execute(f"""
                SELECT material_name, primary_barcode::TEXT
                FROM {tbl}
                WHERE UPPER(source_customer) = UPPER(%s)
                  AND material_name IS NOT NULL AND TRIM(material_name) <> ''
            """, (key,))
            rows = cur.fetchall()
        seen = 0
        for name, bc in rows:
            dv = name.strip()
            if not dv:
                continue
            norm_key = dv.lower()
            bc_clean = str(bc).strip() if bc else None
            if bc_clean in ('0', 'None', ''):
                bc_clean = None
            if norm_key in all_desc:
                all_desc[norm_key]["customers"].add(key)
                if bc_clean and not all_desc[norm_key].get("barcode"):
                    all_desc[norm_key]["barcode"] = bc_clean
            else:
                all_desc[norm_key] = {"customers": {key}, "raw": dv, "barcode": bc_clean}
            seen += 1
        log.info(f"    {key}: {seen:,} rows")
    log.info(f"  Total distinct customer SKUs (across {len(keys)} customers): {len(all_desc):,}")
    return all_desc


# =====================================================================
# Matching (no billing bridge, no DB writes — pure in-memory -> CSV)
# =====================================================================
def build_ai_model():
    log.info(f"  Loading AI model: {CONFIG['model_name']}...")
    model = SentenceTransformer(CONFIG["model_name"])
    log.info(f"  Model loaded (dim={model.get_sentence_embedding_dimension()})")
    return model


def build_faiss_index(model, norm_list: list, cache_key=None):
    cache_dir  = Path(CONFIG["cache_dir"])
    embeddings = None
    if cache_key:
        cache_file = cache_dir / f"embeddings_{cache_key}.pkl"
        if cache_file.exists():
            try:
                cached = pickle.load(open(cache_file, "rb"))
                if cached["count"] == len(norm_list):
                    embeddings = cached["embeddings"]
                    log.info(f"  Cache hit [{cache_key}]")
            except Exception as e:
                log.warning(f"  Cache error: {e}")
    if embeddings is None:
        log.info(f"  Encoding {len(norm_list):,} [{cache_key or 'no-cache'}]...")
        embeddings = model.encode(norm_list, batch_size=CONFIG["embedding_batch_size"],
                                  show_progress_bar=True, normalize_embeddings=True)
        if cache_key:
            cache_dir.mkdir(parents=True, exist_ok=True)
            pickle.dump({"count": len(norm_list), "embeddings": embeddings,
                         "ts": dt.datetime.now().isoformat()},
                        open(cache_dir / f"embeddings_{cache_key}.pkl", "wb"))
    dim   = model.get_sentence_embedding_dimension()
    index = faiss.IndexFlatIP(dim)
    index.add(embeddings.astype(np.float32))
    return index


def _make_result(raw, norm_key, customers, cc, mid, mdesc, score, method, status, bstat) -> dict:
    return {
        "raw_description": raw, "normalized_desc": norm_key,
        "sold_by_customers": customers, "customer_count": cc,
        "matched_master_id": mid, "matched_description": mdesc,
        "similarity_score": score, "match_method": method,
        "match_status": status, "brand_match_status": bstat,
        "classification": None,
    }


def match_all(unique_descs, master_ids, master_raws, master_norms, barcode_lookup,
              model, known_brands):
    top_k    = CONFIG["faiss_top_k"]
    emb_batch = CONFIG["embedding_batch_size"]
    ai_thr   = CONFIG["ai_threshold_matched"]
    fuzz_thr = CONFIG["brand_fuzzy_threshold"]

    master_norm_to_idx = {n: i for i, n in enumerate(master_norms)}
    full_faiss = build_faiss_index(model, master_norms, cache_key="lulu_master_v1")

    all_keys = list(unique_descs.keys())
    total    = len(all_keys)
    results  = []
    counters = dict(barcode=0, exact=0, ai=0, rf=0, unmatched=0, brand_rejected=0)
    start_ts = dt.datetime.now()

    for batch_start in range(0, total, emb_batch):
        batch_keys  = all_keys[batch_start: batch_start + emb_batch]
        fuzzy_keys, fuzzy_norms = [], []

        for key in batch_keys:
            info    = unique_descs[key]
            q_norm  = normalize_text(key)
            cl      = sorted(info["customers"])
            cc      = len(cl)
            barcode = info.get("barcode")

            if barcode and barcode in barcode_lookup:
                mid   = barcode_lookup[barcode]
                mdesc = master_raws[master_ids.index(mid)] if mid in master_ids else ""
                results.append(_make_result(info["raw"], key, cl, cc, mid, mdesc, 100.0,
                                            "barcode", "MATCHED", "n/a"))
                counters["barcode"] += 1
                continue

            if q_norm in master_norm_to_idx:
                idx = master_norm_to_idx[q_norm]
                results.append(_make_result(info["raw"], key, cl, cc,
                                            master_ids[idx], master_raws[idx], 100.0,
                                            "exact", "MATCHED", "n/a"))
                counters["exact"] += 1
                continue

            fuzzy_keys.append(key)
            fuzzy_norms.append(q_norm)

        if not fuzzy_keys:
            continue

        q_emb = model.encode(fuzzy_norms, batch_size=emb_batch,
                             normalize_embeddings=True, show_progress_bar=False).astype(np.float32)

        for i, key in enumerate(fuzzy_keys):
            info   = unique_descs[key]
            cl     = sorted(info["customers"])
            cc     = len(cl)
            q_norm = fuzzy_norms[i]
            q_vec  = q_emb[i: i + 1]

            k = min(top_k, len(master_ids))
            sc_mat, ix_mat = full_faiss.search(q_vec, k)
            ai_scores_k    = sc_mat[0]
            ai_indices_k   = ix_mat[0]

            best_score = 0.0
            best_idx   = int(ai_indices_k[0])
            best_bstat = "no_candidate"
            for rank_j in range(len(ai_indices_k)):
                c_idx = int(ai_indices_k[rank_j])
                if c_idx < 0 or c_idx >= len(master_norms):
                    continue
                adj, bstat = brand_strict_score(q_norm, master_norms[c_idx],
                                                float(ai_scores_k[rank_j]),
                                                known_brands, fuzz_thr)
                if adj > best_score:
                    best_score = adj; best_idx = c_idx; best_bstat = bstat

            rf_pool = [int(ix) for ix in ai_indices_k[:top_k]
                       if 0 <= int(ix) < len(master_norms)]
            best_rf = 0.0; best_rf_idx = best_idx; best_rf_bstat = best_bstat
            for c_idx in rf_pool:
                rf_raw = rapidfuzz_score(q_norm, master_norms[c_idx])
                adj_rf, bstat_rf = brand_strict_score(q_norm, master_norms[c_idx],
                                                       rf_raw, known_brands, fuzz_thr)
                if adj_rf > best_rf:
                    best_rf = adj_rf; best_rf_idx = c_idx; best_rf_bstat = bstat_rf

            if best_rf_idx == best_idx and best_score > 0 and best_rf > 0:
                final_score = 0.65 * best_score + 0.35 * best_rf
                final_idx = best_idx; final_bstat = best_bstat; final_method = "ai+rf"
            elif best_rf > best_score + 0.05 and best_rf >= 0.82:
                final_score = best_rf; final_idx = best_rf_idx
                final_bstat = best_rf_bstat; final_method = "ai+rf"
            else:
                final_score = best_score; final_idx = best_idx
                final_bstat = best_bstat; final_method = "ai"

            is_rejected = "rejected" in final_bstat
            final_score_100 = round(final_score * 100, 2)

            if is_rejected or final_score < ai_thr:
                status = "UNMATCHED"
                counters["unmatched"] += 1
                if is_rejected:
                    counters["brand_rejected"] += 1
                match_id    = master_ids[final_idx]  if not is_rejected else None
                match_desc  = master_raws[final_idx] if not is_rejected else None
                match_score = final_score_100        if not is_rejected else 0.0
            else:
                status = "MATCHED"
                match_id    = master_ids[final_idx]
                match_desc  = master_raws[final_idx]
                match_score = final_score_100
                if "rf" in final_method:
                    counters["rf"] += 1
                else:
                    counters["ai"] += 1

            results.append(_make_result(info["raw"], key, cl, cc, match_id, match_desc,
                                        match_score, final_method, status, final_bstat))

        processed = batch_start + len(batch_keys)
        elapsed   = (dt.datetime.now() - start_ts).total_seconds()
        rate      = processed / elapsed if elapsed > 0 else 0
        eta       = str(dt.timedelta(seconds=int((total - processed) / rate))) if rate > 0 else "?"
        log.info(f"  {processed:>7,}/{total:,} ({100*processed/total:5.1f}%) | "
                 f"bc={counters['barcode']:,} ex={counters['exact']:,} "
                 f"ai={counters['ai']:,} rf={counters['rf']:,} "
                 f"unmatched={counters['unmatched']:,} brand_rej={counters['brand_rejected']:,} | "
                 f"{rate:.0f}/s ETA={eta}")

    return results, counters


def classify_results(results: list) -> list:
    min_score  = CONFIG["classify_min_score"]
    guaranteed = {"barcode", "exact"}
    for r in results:
        method  = r["match_method"]
        status  = r["match_status"]
        score   = r["similarity_score"] or 0.0
        cc      = r["customer_count"]
        matched = (status == "MATCHED" and (score >= min_score or method in guaranteed))
        if matched:
            cls = "FAST MOVER" if cc >= 3 else ("MODERATE MOVER" if cc == 2 else "SLOW MOVER")
        else:
            cls = "LOST OPPORTUNITY"
        r["classification"] = cls
    return results


def save_csv(results: list):
    import pandas as pd
    out_path = CONFIG["csv_output_path"]
    rows = []
    for r in results:
        rows.append({
            "customer": ",".join(r["sold_by_customers"]),
            "customer_count": r["customer_count"],
            "customer_sku": r["raw_description"],
            "lulu_sku": r["matched_description"] or "",
            "lulu_material_code": r["matched_master_id"] or "",
            "score": r["similarity_score"] or 0.0,
            "match_method": r["match_method"],
            "match_status": r["match_status"],
            "brand_match_status": r["brand_match_status"],
            "classification": r["classification"],
        })
    df = pd.DataFrame(rows)
    df.to_csv(out_path, index=False)
    log.info(f"  CSV saved: {out_path} ({len(df):,} rows)")


def main():
    import shutil
    cache_dir = CONFIG["cache_dir"]
    if os.path.exists(cache_dir):
        shutil.rmtree(cache_dir, ignore_errors=True)
    os.makedirs(cache_dir, exist_ok=True)

    log.info("=" * 70)
    log.info(f"LULU-AS-MASTER SKU MATCHING — {dt.datetime.now():%Y-%m-%d %H:%M:%S}")
    log.info(f"Master customer   : {CONFIG['master_key']}")
    log.info(f"Matching customers: {CONFIG['customer_keys']}")
    log.info(f"ai_threshold      : {CONFIG['ai_threshold_matched']}")
    log.info("=" * 70)

    conn = get_conn()
    try:
        log.info("\n[1/6] Checking distinct source_customer values...")
        distinct_vals = print_distinct_source_customers(conn)
        missing = [k for k in [CONFIG["master_key"]] + CONFIG["customer_keys"]
                   if k.upper() not in {v.upper() for v in distinct_vals if v}]
        if missing:
            log.warning(f"  !! These configured keys were NOT found in source_customer: {missing}")
            log.warning(f"  !! Fix CONFIG['master_key'] / CONFIG['customer_keys'] to match the list above and re-run.")

        log.info("\n[2/6] Loading Lulu master...")
        master_ids, master_raws, master_norms, barcode_lookup = load_master(conn)

        log.info("\n[3/6] Loading customer SKUs (TALABAT, C4, GRAND_MALL, AL_MEERA)...")
        unique_descs = load_customers(conn)

        log.info("\n[4/6] Building dynamic brand list...")
        all_norms    = master_norms + [normalize_text(k) for k in unique_descs.keys()]
        known_brands = auto_extend_brands(all_norms, KNOWN_BRANDS, min_freq=5)

        log.info("\n[5/6] Loading AI model & matching...")
        model = build_ai_model()
        results, counters = match_all(unique_descs, master_ids, master_raws, master_norms,
                                      barcode_lookup, model, known_brands)

        log.info("\n[6/6] Classifying & saving CSV...")
        results = classify_results(results)
        save_csv(results)

        log.info("\n" + "=" * 70)
        log.info("PIPELINE COMPLETE ✓")
        log.info(f"  Barcode:        {counters['barcode']:,}")
        log.info(f"  Exact:          {counters['exact']:,}")
        log.info(f"  AI matched:     {counters['ai']:,}")
        log.info(f"  AI+RF matched:  {counters['rf']:,}")
        log.info(f"  Unmatched:      {counters['unmatched']:,}")
        log.info(f"  Brand rejected: {counters['brand_rejected']:,}")
        log.info("=" * 70)

    except Exception as e:
        log.error(f"Pipeline failed: {e}", exc_info=True)
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    main()