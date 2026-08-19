import psycopg2
import pandas as pd

# =========================
# POSTGRES CONFIG
# =========================
PG = {
    "dbname": "test",
    "user": "",
    "password": "",
    "host": "localhost",
    "port": 5432
}

# =========================
# DB HELPER FUNCTIONS
# =========================
def pg_connect():
    return psycopg2.connect(**PG)

def schema_exists(schema):
    conn = pg_connect()
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM information_schema.schemata WHERE schema_name=%s", (schema,))
    exists = cur.fetchone() is not None
    cur.close()
    conn.close()
    return exists

def create_schema(schema):
    conn = pg_connect()
    cur = conn.cursor()
    cur.execute(f'CREATE SCHEMA IF NOT EXISTS "{schema}"')
    conn.commit()
    cur.close()
    conn.close()

def table_exists(schema, table):
    conn = pg_connect()
    cur = conn.cursor()
    cur.execute("""
        SELECT 1 FROM information_schema.tables
        WHERE table_schema=%s AND table_name=%s
    """, (schema, table))
    exists = cur.fetchone() is not None
    cur.close()
    conn.close()
    return exists

def create_table(schema, table, columns):
    """
    Creates a table with all columns as TEXT
    """
    if not columns:
        print(f"⚠️  No columns provided for {schema}.{table}")
        return
    
    col_defs = [f'"{col}" TEXT' for col in columns]
    ddl = f'CREATE TABLE "{schema}"."{table}" ({", ".join(col_defs)})'
    
    conn = pg_connect()
    cur = conn.cursor()
    cur.execute(ddl)
    conn.commit()
    cur.close()
    conn.close()
    print(f"✅ Created table {schema}.{table} with {len(columns)} columns")

# =========================
# READ EXCEL FILE
# =========================
excel_path = r"C:\Users\HP\Downloads\meta_data.csv"

# Read the Excel file
df = pd.read_csv(excel_path)

print(f"Total rows in CSV: {len(df)}")
print(f"Columns in CSV: {df.columns.tolist()}")
print("\n" + "="*80)

# =========================
# PROCESS EACH ROW
# =========================
for idx, row in df.iterrows():
    print(f"\n{'='*80}")
    print(f"Processing row {idx + 1}/{len(df)}")
    print(f"{'='*80}")
    
    # Get destination schema name
    destination_schema = row.get('DESTINATION_SCHEMA_NAME')
    
    # Get source table name (will be used as destination table name)
    source_table = row.get('SOURCE_TABLE_NAME')
    
    # Get the header column which contains column names
    header_text = row.get('HEADER')
    
    print(f"Destination Schema: {destination_schema}")
    print(f"Source Table Name: {source_table}")
    
    # Skip if essential fields are missing
    if pd.isna(destination_schema) or pd.isna(source_table):
        print(f"⚠️  Skipping row - missing schema or table name")
        continue
    
    # Parse column names from HEADER field
    columns = []
    if pd.notna(header_text):
        # Split by newlines and clean up
        columns = [col.strip() for col in str(header_text).split('\n') if col.strip() and col.strip() != '"']
    
    print(f"Columns found: {len(columns)}")
    if columns:
        print(f"First 5 columns: {columns[:5]}")
        print(f"Last 5 columns: {columns[-5:]}")
    
    if not columns:
        print(f"⚠️  Skipping {destination_schema}.{source_table} - no columns found in HEADER")
        continue
    
    # Clean schema and table names
    schema_name = str(destination_schema).strip()
    table_name = str(source_table).strip()
    
    # Remove .dat or other extensions from table name if present
    if '.' in table_name:
        table_name_clean = table_name.split('.')[0]
        print(f"Cleaned table name: {table_name} -> {table_name_clean}")
        table_name = table_name_clean
    
    # Create schema if not exists
    if not schema_exists(schema_name):
        print(f" Creating schema: {schema_name}")
        create_schema(schema_name)
    else:
        print(f"✓ Schema already exists: {schema_name}")
    
    # Create table if not exists
    if not table_exists(schema_name, table_name):
        create_table(schema_name, table_name, columns)
    else:
        print(f"✓ Table already exists: {schema_name}.{table_name}")

print("\n" + "="*80)
print("✅ All schemas and tables processed successfully!")
print("="*80)