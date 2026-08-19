import pandas as pd
from sqlalchemy import create_engine

# ------------------------------
# Database Config
# ------------------------------
DB_USER = ""
DB_PASSWORD = ""
DB_HOST = "localhost"
DB_PORT = "5432"
DB_NAME = "test"

# ------------------------------
# CSV File Path
# ------------------------------
csv_file = r"C:\Users\HP\Downloads\meta_data.csv"  # path to your CSV
table_name = "etl_metadata"            # table to save in DB

# ------------------------------
# Create SQLAlchemy Engine
# ------------------------------
engine = create_engine(f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}")

# ------------------------------
# Read CSV
# ------------------------------
df = pd.read_csv(csv_file)

# ------------------------------
# Save to PostgreSQL
# ------------------------------
df.to_sql(table_name, engine, schema="stg", if_exists='replace', index=False)

print(f"CSV '{csv_file}' successfully saved to table '{table_name}' in database '{DB_NAME}'")
