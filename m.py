import psycopg2
import csv
import re

# -------------------------------
# 🔹 PostgreSQL Connection
# -------------------------------
conn = psycopg2.connect(
    host="localhost",
    database="",
    user="",
    password="",
    port="5432"
)

conn.autocommit = True
cursor = conn.cursor()

# -------------------------------
# 🔹 Read Metadata File
# -------------------------------
file_path = r"C:\Users\HP\Desktop\meta_data_2_extended1.txt"

tables = []

with open(file_path, newline='', encoding="utf-8") as f:
    reader = csv.DictReader(f, delimiter="\t")
    
    for row in reader:
        schema = row["DESTINATION_SCHEMA"]
        table = row["DESTINATION_TABLE"]
        header = row["HEADER"]
        
        if schema and table and header:
            columns = tuple(col.strip().replace(" ", "_").lower() 
                       for col in header.replace('"','').split(","))
            
            tables.append((schema.lower(), table.lower(), columns))

# Remove duplicates
tables = list(set(tables))

# -------------------------------
# 🔹 Create Schemas & Tables
# -------------------------------
for schema, table, columns in tables:
    
    # Create schema
    create_schema_sql = f"CREATE SCHEMA IF NOT EXISTS {schema};"
    cursor.execute(create_schema_sql)
    
    # Create table
    column_definitions = []

    for col in columns:
        col_clean = re.sub(r'\W+', '_', col)
        column_definitions.append(f"{col_clean} TEXT")

    create_table_sql = f"""
    CREATE TABLE IF NOT EXISTS {schema}.{table} (
        {', '.join(column_definitions)}
    );
    """
    
    print(f"Creating table: {schema}.{table}")
    cursor.execute(create_table_sql)

print("✅ All schemas and tables created successfully!")

cursor.close()
conn.close()