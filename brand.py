"""
Brand standardization script.

Usage:
    python standardize_brands.py INPUT.csv [OUTPUT.xlsx]

Goal: collapse brand-name variants to a single canonical brand.
  - Exact-normalized merges (spacing/punctuation/case, IMP/IMPORT suffixes)
  - "PARENT - CHILD" prefixes -> child brand
  - Fuzzy near-duplicate spellings (algorithmic + curated)

  NOTE: prefix collapse (the old Phase C, e.g. "UNILEVER LIPTON" -> "UNILEVER",
  "ULKER BISKREM" -> "ULKER") has been REMOVED. Umbrella/parent-company brands
  like UNILEVER must never swallow their child brands automatically. The only
  merges that happen now are exact-normalized matches and the fuzzy/curated
  matches below -- nothing is inferred from a shared first word anymore.

  FIX (this version): Phase A (exact-normalized) and Phase B (fuzzy) used to
  run on disjoint pools -- fuzzy matching only ever looked at brands that were
  NOT already part of an exact group. That meant once a brand had even one
  exact-normalized twin (e.g. "AMERICAN GARDEN" + "GEMCO - AMERICAN GARDEN"),
  it became invisible to the fuzzy pass, so a genuine typo variant like
  "AMERICN GARDEN" was never found and merged, even though it clears every
  fuzzy threshold. Same story for "BUSH'S BEST" not linking up if "BUSH'S"
  had already exact-matched something else. Fuzzy matching now runs over
  EVERY normalized brand key (singletons and exact-group representatives
  alike), and when a fuzzy hit occurs, the full membership of every matched
  key is merged together. Exact matches (distance 0) are just a special case
  of this now, so Phase A/B are effectively unified.

Outputs an .xlsx with two sheets: the standardized master and the brand crosswalk.
"""

import sys
import re
from collections import defaultdict

import pandas as pd
from rapidfuzz import fuzz, process
from rapidfuzz.distance import Levenshtein

# ---------------------------------------------------------------------------
# CLI ARGS
# ---------------------------------------------------------------------------
INPUT_PATH = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\HP\Downloads\dim_material_master_categorizedollam.xlsx"
OUTPUT_PATH = sys.argv[2] if len(sys.argv) > 2 else r"C:\Users\HP\Downloads\dim_material_master_qnie.csv"
if not OUTPUT_PATH.lower().endswith('.xlsx'):
    OUTPUT_PATH += '.xlsx'

df = pd.read_excel(INPUT_PATH, engine='openpyxl') if INPUT_PATH.lower().endswith('.xlsx') else pd.read_csv(INPUT_PATH)
df['brand'] = df['brand'].fillna('UNKNOWN').astype(str).str.strip()

SUFFIXES = [' IMP', ' IMPORT', ' EXPORT', ' IMPORTED', ' IMPORTS']


def normalize(b):
    s = b.upper()
    s = s.replace("\u2019", "'")
    for suf in SUFFIXES:
        if s.endswith(suf):
            s = s[: -len(suf)]
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


brand_customers = df.groupby('brand')['source_customer'].apply(lambda x: set(x)).to_dict() \
    if 'source_customer' in df.columns else {b: set() for b in df['brand'].unique()}
unique_brands = list(brand_customers.keys())

# NOTE: the raw 'brand' column is never modified anywhere in this script.
# Every fix below (exact/fuzzy/curated merges and the generic-token
# resolution near the bottom) only ever writes into the separate
# 'brand_standardized' output column.

# ---- strip "PARENT COMPANY - PRODUCT BRAND" prefixes, then bucket by
#      normalized alias. This is Phase A (exact-normalized) grouping. ----
def alias_of(b):
    if ' - ' in b:
        return b.rsplit(' - ', 1)[-1].strip()
    return b


norm_map = defaultdict(list)
for b in unique_brands:
    norm_map[normalize(alias_of(b))].append(b)

exact_groups = {k: v for k, v in norm_map.items() if len(v) > 1}

# ---- Phase B: algorithmic fuzzy auto-groups ----
# Runs over ALL normalized keys (not just ones with no exact twin -- see the
# FIX note in the module docstring). This is what lets "AMERICN GARDEN"
# reach "AMERICANGARDEN" even though the latter already has 2 members.
HARD_EXCLUDE = {
    frozenset(['CAPRI', 'CAPRIO']), frozenset(['BARAKAT', 'BARAKA']),
    frozenset(['BUSAN', 'BUSTAN']), frozenset(['BENATURAL', 'NATURAL']),
    frozenset(['ALWANI', 'HALWANI']),
    # Found while auditing the fixed fuzzy pass below -- these clear the
    # distance/prefix thresholds but are genuinely different brands, not
    # spelling variants, so they're explicitly blocked rather than merged.
    frozenset(['CLASSIC', 'CLASSICO']),      # Classic (generic label) vs Classico (pasta sauce brand)
    frozenset(['GOODWAY', 'GOODDAY']),       # Goodway vs Britannia's "Good Day" biscuits
    frozenset(['WALTERS', 'WALKERS']),       # Walters vs Walker's crisps -- different names
    frozenset(['EUROBAKE', 'EUROCAKE']),     # Eurobake vs Eurocake -- different product words
    frozenset(['SUMMERFRESH', 'SUPERFRESH']),  # Summer Fresh vs Super Fresh -- different names
}
MIN_LEN_FOR_AUTO = 7
all_keys = list(norm_map.keys())
used = set()
fuzzy_groups = []
for key in all_keys:
    if key in used:
        continue
    candidates = process.extract(key, all_keys, scorer=fuzz.ratio, limit=8, score_cutoff=88)
    group_keys = [key]
    for cand_key, score, _ in candidates:
        if cand_key == key or cand_key in used:
            continue
        dist = Levenshtein.distance(key, cand_key)
        maxlen = max(len(key), len(cand_key))
        minlen = min(len(key), len(cand_key))
        same_prefix = key[:2] == cand_key[:2]
        if frozenset([key, cand_key]) in HARD_EXCLUDE:
            continue
        allowed_dist = 1 if maxlen <= 10 else 2
        if same_prefix and minlen >= MIN_LEN_FOR_AUTO and dist <= allowed_dist:
            group_keys.append(cand_key)
    # A "group" worth recording is either a genuine fuzzy merge (>1 key) or
    # an already-multi-member exact group passing through untouched.
    if len(group_keys) > 1 or len(norm_map[key]) > 1:
        members = []
        for k in group_keys:
            members.extend(norm_map[k])
        fuzzy_groups.append(members)
        for k in group_keys:
            used.add(k)

# ---- Curated manual merges (additive on top of the algorithmic pass) ----
CURATED_MERGES = [
    ['DEEMAH', 'DEEMAHH'],
    ['DI MARTINO', 'G.DI MARTINO'],
    ['DIVELA', 'DIVELLA'],
    ['MAJDI', 'AJDI'],
    ['ASAK', 'ASMAK'],
    ['DSCHAR', 'SCHAR'],
    ['HAAGE', 'HAAGEN', 'HAAGENDAZS'],
    ['KASIH', 'KASSIH'],
    ['LOACKER', 'LOCKER'],
    ['LOTTE', 'LOTTEE'],
    ['LOV', 'LOVE'],
    ['MC VIITIES', "MCVITIES'S"],
    ['MUCHEE', 'MUNCHEE'],
    ['NAMET', 'NAMLET'],
    ['OREAO', 'OREO'],
    ['POMI', 'POMMI'],
    ['RAYAN', 'RAYYAN'],
    ['SEARA', 'SEARRA'],
    ['SHAHAD', 'SHAHD'],
    ['SICAM', 'SICA'],
    ['SPRING', 'SPRNG'],
    ['TACHIBO', 'TCHIBO'],
    ['TAYBAT', 'TAYEBAT'],
    ['SAHHA', 'SAHA'],
    ['ALICAFE', 'ALICAFÉ'],
    ['BRITANNIA', 'BRITANIA'],
    ['TIM'],
    ['ULKER', 'PLADIS : ULKER', 'PLADIS : ULKER - BISKREM', 'PLADIS : ULKER - HALLEY', 'PLADIS :ULKER - KAT KAT TAT', 'PLADIS : ULKER KELLOGGS', 'PLADIS : ULKER - RONDO', 'ULKER KELLOGGS'],
    ['AMERICANA GARDEN', 'AMERICAN GARDEN'],
]

CURATED_MERGES += [
    ['3SUNICH', 'SUNICH'],
    ['ABOSHEBA', 'BOSHEBA'],
    ['ACORASA', 'ACORSA', 'ACORS'],
    ['ACQUA PANNA', 'AQUA PANNA'],
    ['AL AILA', 'AL AILAH'],
    ['AL TAGHZIAH', 'TAGHZIAH'],
    ['BANJABI KHORY', 'PANJABI KHORY'],
    ['BINGRE', 'BINGREE'],
    ['BR BEEF', 'BRZ BEEF'],
    ['C LUXURY', 'LUXURY'],
    ['FRITO', 'FRITOS'],
    ['G/SABA', 'SABA'],
    ['GARDEIN', 'GARDEN'],
    ['GDY-PASTA', 'GOODY-PASTA'],
    ['GRAN', 'GRANO'],
    ['HALAAL', 'HALAL'],
    ['HILLI', 'HILLLI'],
    ['HIPO', 'HIPPO'],
    ['AHMAD', 'AHMAD TEA'],
    ['AHMED', 'AHMED FOOD'],
    ['PARLE', 'PARLE - FAB!', 'PARLE - HIDE & SEEK', 'PARLE - KRACKJACK',
     'PARLE - MONACO', 'PARLE - PARLE', 'PARLE - PARLE-G', 'PARLE - POPPINS',
     'PARLE - RUSK', 'PARLE-G'],
]

# New confident merges surfaced by the fixed fuzzy pass -- same-brand
# transliteration/typo variants, reviewed to exclude anything ambiguous.
CURATED_MERGES += [
    ['AL BARAKAH', 'AL BARAKA', 'ALBARAKA'],

    ['AL KABER', 'AL KABEER', 'ALKABEER'],
    ['AL MARRAI', 'AL MARAI', 'ALMARAI'],
    ['AL MARRAI', 'AL MARAAI', 'ALMARAAI'],
    ['AL NUTRICAA', 'AL NUTRICA', 'ALNUTRICA'],
    ['ALKRAMAH', 'AL KARAMAH', 'ALKARAMAH'],
    ['ALWLIMAH', 'AL WALIMAH', 'ALWALIMAH'],
    ['AL TAYAB', 'AL TAYYAB', 'ALTAYYAB'],
]

# ---------------------------------------------------------------------------
# Forced-canonical curated merges
#
# For the groups above, pick_canonical() (shortest/cleanest string, LULU
# preferred) decides the winner. That's wrong for a few real cases where
# the "cleanest" string isn't the real brand name:
#   - "KANNAN" is shorter than "KANNAN DEVAN" but "Kannan Devan" (Kannan
#     Devan Hills Plantations) is the actual tea brand -- "Kannan" alone is
#     a truncated read of it, confirmed by material_name overlap
#     ("KANNAN DEVAN TEA 400GR" filed under bare brand "KANNAN").
#   - "KELLOGG'S - CORN FLAKES" / "- RICE KRISPIES" / "- FROSTIES" etc. use
#     the same " - " pattern as parent/child brands (e.g. UNILEVER - LIPTON)
#     but here the part after the dash is a PRODUCT LINE, not a distinct
#     brand -- material_name confirms every one of these is literally
#     Kellogg's-branded stock, so they collapse to the parent instead of
#     being split apart by alias_of(). "KELLOGG'S - PRINGLES" is
#     deliberately left OUT: its material_name ("PRINGLES HOT KICKIN SOUR
#     CRM CHIPS") never mentions Kellogg's, so it looks like a mis-tagged
#     row rather than a real Kellogg's product -- worth a manual check
#     rather than folding in automatically.
#   - "BUSH" (bare, no apostrophe) was left out of the BUSH'S group earlier
#     as ambiguous, but material_name evidence settles it: "Bush Veg.Baked
#     Bean 235g" is the same product line as "Bush's Pinto Beans" /
#     "Bush's Best Vegetarian Baked Beans" -- same brand, apostrophe just
#     got dropped in this row.
#   - "KHALAS AL NADEEM" and bare "KHALAS" share the identical
#     material_name ("KHALAS AL NADEEM MALKI DATES") -- same product, two
#     different brand-field values -- so they collapse to "KHALAS".
# ---------------------------------------------------------------------------
FORCED_CANONICAL_MERGES = [
    ("BUSH'S", ['BUSH', 'BUSHS', "BUSH'S", "BUSH'S BEST"]),
    ('KANNAN DEVAN', ['KANNAN DEVAN', 'KANNAN']),
    ("KELLOGG'S", [
        "KELLOGG'S", 'KELLOGGS', 'KELLOGS', 'KELLOGGS\'S', 'KELLOG\'S', 'KELLOGG S', 'KELLOGG S IMP',
        "KELLOGG'S - CORN FLAKES", "KELLOGG'S - EGGO", "KELLOGG'S - RICE KRISPIES",
        "KELLOGG'S - FROSTIES", "KELLOGG'S - SPECIAL K", "KELLOGG'S - KELLOGG'S",
        "KELLOGG'S - COCO POPS", "KELLOGG'S - POP TARTS",
        "KELLOGG'S FROOT LOOPS", "KELLOGG'S RICE KRSPIES", "KELLOGG'S CORN POP",
        "KELLOGG'S GRANOLA",
    ]),
    ('KHALAS', ['KHALAS', 'KHALAS AL NADEEM']),
]

fuzzy_groups.extend(CURATED_MERGES)

# ---------------------------------------------------------------------------
# Canonical selection & mapping
# ---------------------------------------------------------------------------
PUNCT = "'.-"


def clean_score(b):
    return (any(c in b for c in PUNCT), len(b))


def pick_canonical(members):
    lulu_members = [m for m in members if 'LULU' in brand_customers.get(m, set())]
    pool = lulu_members if lulu_members else members
    return sorted(pool, key=clean_score)[0]


brand_to_std = {}
group_report = []


def assign(members, canon, match_type):
    for b in members:
        brand_to_std[b] = canon
    custs = sorted(set().union(*[brand_customers.get(b, set()) for b in members])) \
        if members else []
    group_report.append({'members': members, 'standardized': canon,
                         'match_type': match_type, 'customers_involved': custs})


for v in fuzzy_groups:
    v = sorted(set(b for b in v if b in brand_customers))
    if len(v) < 2:
        continue
    match_type = 'exact_normalized' if len({normalize(alias_of(b)) for b in v}) == 1 else 'fuzzy_curated'
    assign(v, pick_canonical(v), match_type)

# Forced-canonical merges run last so they win over whatever the automatic
# passes above decided for the same raw brands.
for canon, members in FORCED_CANONICAL_MERGES:
    v = sorted(set(b for b in members if b in brand_customers))
    if len(v) < 2:
        continue
    assign(v, canon, 'curated_forced_canonical')

for b in unique_brands:
    brand_to_std.setdefault(b, b)

df['brand_standardized'] = df['brand'].map(brand_to_std)

# ---------------------------------------------------------------------------
# Phase D: Row-level resolution of generic/truncated brand tokens
#
# Some rows only ever captured "AL" as the brand (a truncated/generic
# value), while the real brand -- "AL ARZ", "AL BADAL", "AL AILA", etc. --
# is embedded at the start of material_name. For rows whose raw brand is
# exactly one of GENERIC_BRAND_TOKENS, take the first two words of
# material_name (falling back to the first word) and check whether that
# string already exists elsewhere in the data as a real brand. If so, set
# brand_standardized for that row to the (fully standardized) real brand.
#
# IMPORTANT: this only ever writes to df['brand_standardized']. The raw
# 'brand' column is left completely untouched, per request.
# ---------------------------------------------------------------------------
GENERIC_BRAND_TOKENS = {'AL'}  # add more bare/truncated tokens here if found

if 'material_name' in df.columns:
    existing_brand_norms = {normalize(b): b for b in unique_brands}
    generic_mask = df['brand'].astype(str).str.strip().str.upper().isin(GENERIC_BRAND_TOKENS)

    def resolve_generic_brand(row):
        name = str(row['material_desc']).strip()
        words = name.split()
        if not words:
            return row['brand_standardized']
        raw_brand_norm = normalize(str(row['brand']))
        # try the 2-word prefix first (e.g. "AL ARZ", "AL BADAL"),
        # then fall back to the 1-word prefix
        for n in (2, 1):
            if len(words) >= n:
                candidate = ' '.join(words[:n])
                cand_norm = normalize(candidate)
                if cand_norm in existing_brand_norms and cand_norm != raw_brand_norm:
                    matched_raw = existing_brand_norms[cand_norm]
                    return brand_to_std.get(matched_raw, matched_raw)
        return row['brand_standardized']

    df.loc[generic_mask, 'brand_standardized'] = df.loc[generic_mask].apply(resolve_generic_brand, axis=1)
    still_generic = int((df.loc[generic_mask, 'brand_standardized']
                         .astype(str).str.strip().str.upper().isin(GENERIC_BRAND_TOKENS)).sum())
    print(f'Rows still stuck on a generic brand token after resolution: {still_generic}')

# ---------------------------------------------------------------------------
# Build crosswalk table
# ---------------------------------------------------------------------------
examples = df.groupby('brand')['material_desc'].apply(
    lambda x: ' | '.join(x.dropna().astype(str).unique()[:3])).to_dict() \
    if 'material_desc' in df.columns else {}
counts = df.groupby('brand').size().to_dict()

rows = []
for g in group_report:
    matched_to_lulu = 'LULU' in g['customers_involved']
    for m in g['members']:
        rows.append({
            'raw_brand': m,
            'standardized_brand': g['standardized'],
            'match_type': g['match_type'],
            'customers_involved': '|'.join(g['customers_involved']),
            'matched_to_lulu': matched_to_lulu,
            'row_count': counts.get(m, 0),
            'sample_material_names': examples.get(m, ''),
        })
# keep='last': later groups (curated, then forced-canonical) are applied
# after earlier automatic ones and are meant to win, so the report should
# reflect the LAST group each raw brand ended up in, not the first --
# otherwise the crosswalk can show a stale pre-override standardization
# even though brand_to_std / brand_standardized already has it right.
crosswalk = pd.DataFrame(rows).drop_duplicates(subset=['raw_brand'], keep='last')

# ---------------------------------------------------------------------------
# Write Excel (two sheets)
# ---------------------------------------------------------------------------
with pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='standardized_master', index=False)
    crosswalk.to_excel(writer, sheet_name='brand_crosswalk', index=False)

    from openpyxl.styles import Font
    for ws in writer.book.worksheets:
        for cell in ws[1]:
            cell.font = Font(name='Arial', bold=True)
        for col in ws.columns:
            width = min(max((len(str(c.value)) for c in col if c.value is not None), default=10) + 2, 60)
            ws.column_dimensions[col[0].column_letter].width = width

print('Input:', INPUT_PATH)
print('Output:', OUTPUT_PATH)
print('Groups formed (exact + fuzzy + curated):', len(group_report))
print('Rows with brand changed:', int((df['brand'] != df['brand_standardized']).sum()))