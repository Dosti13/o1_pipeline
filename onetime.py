import pandas as pd
import logging
import os
import dotenv 
from pathlib import Path
from sqlalchemy import create_engine, text
from sqlalchemy.types import Text
dotenv.load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
log = logging.getLogger("source_to_stg")

TARGET_SCHEMA = "stg"

# source file -> destination table
SOURCE_TABLE_MAP = {
    r"C:\Users\HP\Downloads\AL_MEERA_FILTERED.xlsx":               "scrap_almeera",
    r"C:\Users\HP\Downloads\C4_FILTERED.xlsx":                     "scrap_c4",
    r"C:\Users\HP\Downloads\LULU_FILTERED.xlsx":                   "scrap_lulu",
    r"C:\Users\HP\Downloads\TALABAT_ALL_LOCATIONS_FILTERED.xlsx":  "scrap_talabat",
    r"C:\Users\HP\Downloads\GRAND_MALL.csv":                       "scrap_grand_mall",
    r"C:\Users\HP\Downloads\SPAR_FILTERED.xlsx":                   "scrap_spar",
}

ENGINE_URL = f"postgresql+psycopg2://{os.getenv('PG_USER')}:{os.getenv('PG_PASSWORD')}@{os.getenv('PG_HOST')}:5432/{os.getenv('PG_DATABASE')}"


def read_source(path: str) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path, sheet_name=0)


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    # Convert everything to string but keep real NaN as NULL
    # (astype(str) alone turns NaN into the literal string "nan")
    df = df.applymap(lambda x: str(x) if pd.notna(x) else None)
    return df


def truncate_table(engine, schema: str, table: str):
    with engine.begin() as conn:
        conn.execute(text(f'TRUNCATE TABLE "{schema}"."{table}";'))
    log.info("Truncated %s.%s", schema, table)


def load_file(engine, path: str, table: str):
    if not Path(path).exists():
        log.error("File not found, skipping: %s", path)
        return

    log.info("Reading %s", path)
    df = read_source(path)
    df = clean_dataframe(df)

    truncate_table(engine, TARGET_SCHEMA, table)

    dtype_dict = {col: Text() for col in df.columns}
    df.to_sql(
        name=table,
        schema=TARGET_SCHEMA,
        con=engine,
        if_exists="append",   # table already truncated above, so this is effectively a full replace
        index=False,
        dtype=dtype_dict,
    )
    log.info("Loaded %d rows into %s.%s", len(df), TARGET_SCHEMA, table)


def run():
    engine = create_engine(ENGINE_URL)
    for path, table in SOURCE_TABLE_MAP.items():
        try:
            load_file(engine, path, table)
        except Exception as e:
            log.error("Failed loading %s -> %s.%s: %s", path, TARGET_SCHEMA, table, e)


if __name__ == "__main__":
    run()