"""
Pipeline: Excel -> Postgres (append)

Reads the final merged Excel file (material_master + needs_review/category_method
pulled in from the second file) and appends it into:

    database : qnie_new_test
    schema   : sales
    table    : material_master_cx

Run:
    python load_to_postgres.py --dry-run   # preview only
    python load_to_postgres.py             # actually append
"""

import argparse
import logging
import sys

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy import types as satypes

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------

DB_NAME = ""
DB_SCHEMA = ""
DB_TABLE = ""

# Path to the Excel file to load. Set this directly here, or override it with --file
# when running the script (command line value always wins if provided).
FILE_PATH = r"C:\Users\HP\Downloads\dim_material_master_qnie_csv_RESOLVED_v3_updated_final.xlsx"
SHEET_NAME = 0

# Explicit column types matching the actual Postgres table schema.
# This avoids pandas/SQLAlchemy failing to infer a type for columns that are
# all-None (e.g. similarity, customer_category) or mixed-type.
SQL_DTYPES = {
    "material_code":     satypes.Text(),
    "material_name":     satypes.Text(),
    "brand":              satypes.Text(),
    "source_customer":    satypes.Text(),
    "customer_category":  satypes.Text(),
    "lulu_category":      satypes.Text(),
    "confidence":         satypes.Float(),
    "similarity":         satypes.Float(),
    "method":             satypes.Text(),
    "review":             satypes.Float(),
    "barcode":            satypes.Text(),
}

# Column mapping: DB_COLUMN -> SOURCE_COLUMN (in your Excel file)
# Edit the right-hand side if a source column name differs.
# Set the value to None for columns that don't exist in the source yet —
# they'll be created as empty/NULL so the load doesn't break.
COLUMN_MAP = {
    "material_code":     "material_code",
    "material_name":     "material_desc",
    "brand":              "brand",
    "source_customer":    None,   # constant value below (CONSTANT_VALUES), not pulled from a column
    "customer_category":  None,   # TODO: not present in current file — fill in source col or leave NULL
    "lulu_category":      "lulu_category_final",
    "confidence":         "similarity_confidence_score",
    "similarity":         None,   # TODO: confirm if this should be a separate score column
    "method":             "category_method",   # comes from the joined "other file"
    "review":             "needs_review",      # comes from the joined "other file" — confirm exact spelling in your file!
    "barcode":            "barcode",
}

# Fixed/constant values to stamp onto every row (overrides COLUMN_MAP for these DB columns).
CONSTANT_VALUES = {
    "source_customer": "QNIE",
}

CHUNKSIZE = 5000  # rows per batch insert

# --------------------------------------------------------------------------
# LOGGING
# --------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)


# --------------------------------------------------------------------------
# DB CONNECTION
# --------------------------------------------------------------------------

def get_engine():
    """
    NOTE: credentials are hardcoded here for local testing convenience.
    Before sharing this script or committing it anywhere, move these back to
    environment variables (PGHOST, PGUSER, PGPASSWORD, PGDATABASE) so the
    password isn't sitting in a plain-text file.
    """
    host = ""
    port = "5432"
    user = ""
    password = ""
    database = DB_NAME

    conn_str = f"postgresql+psycopg2://{user}:{password}@{host}:{port}/{database}"
    return create_engine(conn_str)


# --------------------------------------------------------------------------
# LOAD + TRANSFORM
# --------------------------------------------------------------------------

def load_excel(filepath: str, sheet_name=0) -> pd.DataFrame:
    log.info(f"Reading {filepath} (sheet={sheet_name})")
    df = pd.read_excel(filepath, sheet_name=sheet_name)
    log.info(f"Loaded {len(df)} rows, {len(df.columns)} columns")
    return df


def build_output_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Selects/renames columns per COLUMN_MAP, producing exactly the DB's column set."""
    out = pd.DataFrame()
    missing_cols = []

    for db_col, src_col in COLUMN_MAP.items():
        if db_col in CONSTANT_VALUES:
            out[db_col] = CONSTANT_VALUES[db_col]
            continue
        if src_col is None:
            out[db_col] = None
            continue
        if src_col not in df.columns:
            missing_cols.append((db_col, src_col))
            out[db_col] = None
            continue
        out[db_col] = df[src_col]

    if missing_cols:
        log.warning("Some mapped source columns were not found in the file:")
        for db_col, src_col in missing_cols:
            log.warning(f"  {db_col} <- '{src_col}' (not found, filled with NULL)")

    # Basic cleanup
    out["material_code"] = out["material_code"].astype(str).str.strip()
    out = out.dropna(subset=["material_code"])
    out = out[out["material_code"] != ""]

    # review column in the table is double precision, but the source often has
    # True/False (or NaN) — convert booleans to 1.0/0.0 so it matches the target type.
    if "review" in out.columns:
        out["review"] = out["review"].map({True: 1.0, False: 0.0, "True": 1.0, "False": 0.0}).fillna(out["review"])
        out["review"] = pd.to_numeric(out["review"], errors="coerce")

    # confidence / similarity must be numeric (double precision in the table)
    for numeric_col in ("confidence", "similarity"):
        if numeric_col in out.columns:
            out[numeric_col] = pd.to_numeric(out[numeric_col], errors="coerce")

    return out


# --------------------------------------------------------------------------
# LOAD TO POSTGRES
# --------------------------------------------------------------------------

def append_to_postgres(df: pd.DataFrame, engine, dry_run: bool = False):
    log.info(f"Prepared {len(df)} rows for table {DB_SCHEMA}.{DB_TABLE}")

    if dry_run:
        log.info("Dry run enabled — no data will be written. Preview:")
        print(df.head(10).to_string())
        print(f"\nTotal rows that would be appended: {len(df)}")
        return

    with engine.begin() as conn:
        # sanity check the table exists
        check = conn.execute(text("""
            SELECT to_regclass(:full_table)
        """), {"full_table": f"{DB_SCHEMA}.{DB_TABLE}"}).scalar()
        if check is None:
            log.error(f"Table {DB_SCHEMA}.{DB_TABLE} does not exist. Aborting.")
            sys.exit(1)

    df.to_sql(
        name=DB_TABLE,
        con=engine,
        schema=DB_SCHEMA,
        if_exists="append",
        index=False,
        chunksize=CHUNKSIZE,
        method="multi",
        dtype=SQL_DTYPES,
    )
    log.info(f"Appended {len(df)} rows into {DB_SCHEMA}.{DB_TABLE}")


# --------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Load merged material master file into Postgres.")
    parser.add_argument("--file", default=FILE_PATH, help=f"Path to the merged Excel file (default: {FILE_PATH})")
    parser.add_argument("--sheet", default=SHEET_NAME, help=f"Sheet name or index (default: {SHEET_NAME})")
    parser.add_argument("--dry-run", action="store_true", help="Preview only, don't write to DB")
    args = parser.parse_args()

    df = load_excel(args.file, sheet_name=args.sheet)
    out = build_output_frame(df)

    if args.dry_run:
        append_to_postgres(out, engine=None, dry_run=True)
        return

    engine = get_engine()
    append_to_postgres(out, engine, dry_run=False)


if __name__ == "__main__":
    main()