import pandas as pd
import logging
from sqlalchemy import create_engine
import os
import dotenv

dotenv.load_dotenv()

# Logging configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

logging.info("Starting pipeline")

INPUT_FILE = r"C:\Users\HP\Desktop\zero1\Exact Matches.csv"
INPUT_SHEET = 0  # set to None if reading a plain CSV

engine = None

try:
    # --------------------------------------------------
    # 1. Read input file
    # --------------------------------------------------
    logging.info(f"Reading input file: {INPUT_FILE}")

    if INPUT_FILE.endswith(".xlsx") or INPUT_FILE.endswith(".xls"):
        df = pd.read_excel(
            INPUT_FILE,
            sheet_name=INPUT_SHEET
        )
    else:
        df = pd.read_csv(INPUT_FILE)

    logging.info(f"File read successfully. Rows: {len(df)}, Columns: {len(df.columns)}")

    # --------------------------------------------------
    # 2. Clean column names
    # --------------------------------------------------
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )

    logging.info("Column names cleaned")

    # --------------------------------------------------
    # 3. Convert values to strings while preserving NULL
    # --------------------------------------------------
    df = df.applymap(
        lambda x: str(x) if pd.notna(x) else None
    )

    logging.info("Data type conversion completed")

    print(df.head())

    # --------------------------------------------------
    # 4. PostgreSQL connection
    # --------------------------------------------------
    logging.info("Creating PostgreSQL connection")

    engine = create_engine(
        f"postgresql+psycopg2://"
        f"{os.getenv('PG_USER')}:"
        f"{os.getenv('PG_PASSWORD')}@"
        f"{os.getenv('PG_HOST')}:5432/"
        f"{os.getenv('PG_DB')}"
    )

    logging.info("PostgreSQL engine created successfully")

    # --------------------------------------------------
    # 5. Load data into PostgreSQL
    # --------------------------------------------------
    logging.info("Loading data into stg.scrap_talabat")

    df.to_sql(
        name="scrap_talabat",
        schema="stg",
        con=engine,
        if_exists="replace",
        index=False,
    )

    logging.info(
        f"Data loaded successfully. "
        f"{len(df)} rows inserted into stg.scrap_talabat"
    )

except FileNotFoundError:
    logging.error(
        f"Input file not found: {INPUT_FILE}",
        exc_info=True
    )

except pd.errors.EmptyDataError:
    logging.error(
        "Input file is empty",
        exc_info=True
    )

except pd.errors.ParserError:
    logging.error(
        "Error while parsing the input file",
        exc_info=True
    )

except Exception as e:
    logging.error(
        f"Pipeline failed: {e}",
        exc_info=True
    )

finally:
    # --------------------------------------------------
    # 6. Close database connection
    # --------------------------------------------------
    if engine is not None:
        engine.dispose()
        logging.info("PostgreSQL connection closed")

    logging.info("Pipeline finished")