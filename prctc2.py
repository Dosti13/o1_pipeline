import os

import pandas as pd
import logging
from sqlalchemy import create_engine
from sqlalchemy.types import Text
import dotenv
dotenv.load_dotenv()
logging.basicConfig(level=logging.INFO)
logging.info("start pipeline")

INPUT_FILE = r"C:\Users\HP\Downloads\almeera_food_products.csv"

# Read file
df = pd.read_csv(INPUT_FILE, dtype={"barcode": str})
# Normalize column names
df.columns = (
    df.columns
      .str.strip()
      .str.lower()
      .str.replace(" ", "_")
)

# Clean barcode
if "barcode" in df.columns:
    df["barcode"] = (
        df["barcode"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)   # Remove trailing .0
    )

# Convert remaining columns to string
for col in df.columns:
    if col != "barcode":
        df[col] = df[col].fillna("").astype(str)

# PostgreSQL connection
engine = create_engine(
   f"postgresql+psycopg2://{os.getenv('PG_USER')}:{os.getenv('PG_PASSWORD')}@{os.getenv('PG_HOST')}:5432/{os.getenv('PG_DATABASE')}"

)

# Store all columns as TEXT
dtype_dict = {col: Text() for col in df.columns}

df.to_sql(
    name="scrap_almeera",
    schema="stg",
    con=engine,
    if_exists="replace",
    index=False,
    dtype=dtype_dict
)

logging.info("Data loaded successfully")