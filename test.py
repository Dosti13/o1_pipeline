import os
import re
import ftplib
import logging
import tempfile
import smtplib
import traceback
import time
from datetime import datetime
from typing import Optional, List, Dict
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pyspark.sql import SparkSession, DataFrame
from pyspark.sql.functions import col, lit, current_timestamp
import psycopg2
from psycopg2 import extras
import dotenv 
import os
dotenv.load_dotenv()

# ---------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------


# ---------------------------------------------------------
# LOCAL DB CONNETION 
# ---------------------------------------------------------
CONFIG = {

    "ftp_host": os.getenv("FTP_HOST"),
    "ftp_user": os.getenv("FTP_USER"),
    "ftp_password": os.getenv("FTP_PASSWORD"),
    "ftp_remote_dir": os.getenv("FTP_DIR"),
    "pg_dbname": os.getenv("PG_DB_LOCAL"),
    "pg_host": os.getenv("PG_HOST_LOCAL","localhost"),
    "pg_port": 5432,
    "pg_user":os.getenv("PG_USER_LOCAL") ,
    "pg_password": os.getenv("PG_PASSWORD_LOCAL"),
    "pg_jdbc_url": f"jdbc:postgresql://localhost:5432/{os.getenv("PG_DB_LOCAL")}",
    "pg_jdbc_driver": "org.postgresql.Driver",
    "encoding": "latin1",
    "global_default_delimiter": "~",
    "extension_delimiters": {".csv": ",", ".xlsx": None, ".xls": None, ".dat": "~", ".txt": ","},
    "moc_pattern": r"MOC_STOCK_\d{8}_\d{6}\.txt$",

# Email Settings
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 465,
    "smtp_user": "dostikhan41@gmail.com",
    "smtp_password": "pqtx cbze alve qfdo",
    "email_from": "dostikhan41@gmail.com",
    "email_to": "dostikhan338@gmail.com",
    "email_cc": "dostmuhammadkhanrind@gmail.com",
}


# ------------------------------------ ---------------------
# LOGGING SETUP
# --------------------------------- ------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.StreamHandler(), logging.FileHandler("etl_job.log", mode='a')]
)

# ---------------------------------------------------------
# EMAIL NOTIFICATION
# ---------------------------------------------------------
def send_etl_summary_email(subject: str, summary_report: str):
    try:
        msg = MIMEMultipart()
        msg['From'] = CONFIG["email_from"]
        msg['To'] = CONFIG["email_to"]
        if CONFIG["email_cc"]: msg['Cc'] = CONFIG["email_cc"]
        msg['Subject'] = subject
        body = f"Hello,\n\nETL Job execution finished. Summary report:\n\n{summary_report}\n\nRegards,\nETL System"
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        recipients = [CONFIG["email_to"]]
        if CONFIG["email_cc"]: recipients.extend([r.strip() for r in CONFIG["email_cc"].split(',')])
        with smtplib.SMTP_SSL(CONFIG["smtp_server"], CONFIG["smtp_port"]) as server:
            server.login(CONFIG["smtp_user"], CONFIG["smtp_password"])
            server.sendmail(CONFIG["email_from"], recipients, msg.as_string())
        logging.info("Summary email sent successfully.")
    except Exception as e:
        logging.error(f"CRITICAL: Failed to send email: {e}")

# ---------------------------------------------------------
# DATABASE UTILITIES & TRACKING
# ---------------------------------------------------------
def pg_connect():
    return psycopg2.connect(
        dbname=CONFIG["pg_dbname"], user=CONFIG["pg_user"],
        password=CONFIG["pg_password"], host=CONFIG["pg_host"], port=CONFIG["pg_port"]
    )

def log_etl_tracking(source_table, dest_schema, dest_table, file_name, start_time, status, records=0, error=""):
    """Saves execution history to stg_test.etl_job_tracking."""
    try:
        conn = pg_connect()
        cur = conn.cursor()
        query = """
            INSERT INTO stg_test.etl_job_tracking 
            (source_table_name, destination_schema, destination_table, file_name, start_time, end_time, status, records_processed, error_message)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        cur.execute(query, (source_table, dest_schema, dest_table, file_name, start_time, datetime.now(), status, records, str(error)[:500]))
        conn.commit()
        cur.close(); conn.close()
    except Exception as e:
        logging.error(f"Tracking Database Update Failed: {e}")

def check_table_exists(schema: str, table: str) -> bool:
    try:
        conn = pg_connect(); cur = conn.cursor()
        cur.execute("SELECT 1 FROM information_schema.tables WHERE table_schema = %s AND table_name = %s", (schema, table))
        exists = cur.fetchone() is not None
        cur.close(); conn.close()
        return exists
    except Exception as e:
        return False

def get_table_columns(schema: str, table: str) -> List[str]:
    try:
        conn = pg_connect(); cur = conn.cursor()
        cur.execute("SELECT column_name FROM information_schema.columns WHERE table_schema = %s AND table_name = %s ORDER BY ordinal_position", (schema, table))
        columns = [r[0] for r in cur.fetchall()]
        cur.close(); conn.close()
        return columns
    except Exception as e:
        return []

# ---------------------------------------------------------
# LOAD METADATA
# ---------------------------------------------------------
def load_metadata() -> List[Dict]:
    rows = []
    try:
        conn = pg_connect()
        query = """
            SELECT * FROM stg.etl_metadata 
            WHERE CASE WHEN "LOAD_FREQUENCY"::text ~ '^[0-9]+$' THEN "LOAD_FREQUENCY"::int ELSE 0 END >= 0
        """
        cur = conn.cursor(cursor_factory=extras.RealDictCursor)
        cur.execute(query)

        for meta in cur.fetchall():
            print(meta, "\n-------------------  -------------------")
            if 'where' in meta and 'where_clause' not in meta:
                meta['where_clause'] = meta.pop('where')
            
            header_raw = meta.get('HEADER')
            print(header_raw," <<<<<<<<<<<<<<<<<<<<<<<<<< HEADER RAW")
            if header_raw and isinstance(header_raw, str):
                cleaned = str(header_raw).strip()
                if cleaned.upper() in ['FULL', '0', '', 'NULL', 'NONE']:
                    meta['headers_list'] = None
                else:
                    cleaned = cleaned.replace('\\', '\n').replace('\r\n', '\n').replace('\r', '\n').replace(',', '\n')
                    meta['headers_list'] = [h.strip() for h in cleaned.split('\n') if h.strip()]
            else:
                meta['headers_list'] = None

            load_table_val = str(meta.get('LOAD_TABLE') or meta.get('LOAD_TYPE') or '').strip().upper()
            print(load_table_val,"\n <<<<<<<<<<<<<<<<<<<<<<<<<< LOAD TABLE VALUE \n")
            meta['write_mode'] = 'overwrite' if 'TRUNCATE' in load_table_val else 'append'
            rows.append(meta)
        cur.close(); conn.close()
    except Exception as e:
        logging.error(f"Metadata Load Failed: {e}")
    return rows

# ---------------------------------------------------------
# FILE HELPERS
# ---------------------------------------------------------
def resolve_delimiter(meta: Dict, filename: str) -> str:
    delim = meta.get('file_delimiter') or meta.get('delimiter')
    if delim and str(delim).strip(): return str(delim).strip()
    src = meta.get('SOURCE_TABLE_NAME', '').lower()
    _, ext = os.path.splitext(filename)
    if src == 'mocstock' and ext.lower() == '.txt': return ","
    return CONFIG["extension_delimiters"].get(ext.lower(), CONFIG["global_default_delimiter"])

def download_file(remote_filename: str) -> Optional[str]:
    suffix = os.path.splitext(remote_filename)[1] or ".csv"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp_path = tmp.name
    tmp.close()
    try:
        ftp = ftplib.FTP(CONFIG["ftp_host"])
        ftp.login(CONFIG["ftp_user"], CONFIG["ftp_password"])
        ftp.cwd(CONFIG["ftp_remote_dir"])
        with open(tmp_path, "wb") as f:
            ftp.retrbinary(f"RETR " + remote_filename, f.write)
        ftp.quit() 
        return tmp_path
    except Exception as e:
        if os.path.exists(tmp_path): os.remove(tmp_path)
        return None

def align_schema(df: DataFrame, dest_schema: str, dest_table: str) -> DataFrame:
    target_cols = get_table_columns(dest_schema, dest_table)
    if not target_cols: return df
    df_cols_lower = {c.lower(): c for c in df.columns}
    for target_col in target_cols:
        if target_col.lower() not in df_cols_lower:
            df = df.withColumn(target_col, lit(None).cast("string"))
        else:
            actual_col = df_cols_lower[target_col.lower()]
            if actual_col != target_col: df = df.withColumnRenamed(actual_col, target_col)
    return df.select(*target_cols)

# ------------------------------------------------------- --
# CORE PROCESSOR
# ------------------------------------------------------- --
def process_metadata_row(meta: Dict) -> str:
    start_time = datetime.now()
    src_table_original = str(meta.get('SOURCE_TABLE_NAME') or "").strip()
    dest_schema = str(meta.get('DESTINATION_SCHEMA_NAME') or "stg_test").strip()
    dest_table = str(meta.get('DESTINATION_DATABASE') or meta.get('destination') or src_table_original).strip()
    local_path = None
    target_file = "N/A"

    try:
        logging.info(f"--- Processing: {src_table_original} -> {dest_schema}.{dest_table} ---")
        
        # FTP File Discovery
        ftp = ftplib.FTP(CONFIG["ftp_host"])
        ftp.login(CONFIG["ftp_user"], CONFIG["ftp_password"])
        ftp.cwd(CONFIG["ftp_remote_dir"])
        ftp_files = ftp.nlst()
        ftp.quit()

        if src_table_original.lower() == "mocstock":
            matches = [f for f in ftp_files if re.search(CONFIG["moc_pattern"], f, re.IGNORECASE)]
            if matches: target_file = sorted(matches)[-1]
        else:
            possible_names = [src_table_original, f"{src_table_original}.csv", f"{src_table_original}.dat", f"{src_table_original}.txt"]
            for name in possible_names:
                if name in ftp_files: target_file = name; break

        if target_file == "N/A":
            raise FileNotFoundError(f"Source file {src_table_original} not found on FTP.")

        # Download
        local_path = download_file(target_file)
        if not local_path: raise Exception("Download failed.")

        # Read
        delimiter = resolve_delimiter(meta, target_file)
        headers_list = meta.get('headers_list')
        ext = os.path.splitext(target_file)[1].lower()
        
        if ext in [".xlsx", ".xls"]:
            df = spark.read.format("com.crealytics.spark.excel").option("header", "true").load(local_path)
        else:
            # Header false read karenge taake manual count validation ho sake
            df = spark.read.option("header", "false" if headers_list else "true").option("delimiter", delimiter).csv(local_path)

        # --- UPDATED COLUMN COUNT VALIDATION LOGIC ---
        if headers_list:
            # load_current_timestamp metadata mein hota hai par file mein nahi hota, isay count se bahar nikaalo
            file_headers_expected = [h for h in headers_list if h.lower() != 'load_current_timestamp']
            
            source_col_count = len(df.columns)
            expected_col_count = len(file_headers_expected)

            # Strict Comparison
            if source_col_count != expected_col_count:
                # Detail ke liye sample row capture karein
                sample = df.limit(1).collect()
                sample_data = str(sample[0].asDict()) if sample else "Empty File"
                raise ValueError(
                    f"COLUMN COUNT MISMATCH: File '{target_file}' has {source_col_count} columns, "
                    f"but Metadata expects {expected_col_count} data columns (excluding timestamp). "
                    f"Sample Data: {sample_data}"
                )
            
            # Agar count sahi hai, to headers apply karein
            df = df.toDF(*file_headers_expected)
        # ---------------------------------------------

        # Filter & Align
        where_clause = meta.get('where_clause')
        if where_clause and str(where_clause).strip().upper() not in ['', 'FULL', '0', 'NONE']:
            df = df.filter(where_clause)

        if check_table_exists(dest_schema, dest_table):
            df = align_schema(df, dest_schema, dest_table)
        
        df = df.withColumn("load_current_timestamp", current_timestamp())

        # Write
        if meta.get('write_mode') == 'overwrite':
            conn = pg_connect(); cur = conn.cursor()
            cur.execute(f'TRUNCATE TABLE "{dest_schema}"."{dest_table}"')
            conn.commit(); cur.close(); conn.close()

        jdbc_opts = {
            "url": CONFIG["pg_jdbc_url"], "dbtable": f'"{dest_schema}"."{dest_table}"',
            "user": CONFIG["pg_user"], "password": CONFIG["pg_password"],
            "driver": CONFIG["pg_jdbc_driver"], "stringtype": "unspecified"
        }
        df.write.mode("append").format("jdbc").options(**jdbc_opts).save()
        
        count = df.count()
        
        log_etl_tracking(src_table_original, dest_schema, dest_table, target_file, start_time, "SUCCESS", records=count)
        return f"[SUCCESS] {src_table_original}: Loaded {count} rows into {dest_schema}.{dest_table} "

    except Exception as e:
        error_msg = str(e)
        log_etl_tracking(src_table_original, dest_schema, dest_table, target_file, start_time, "FAILED", error=error_msg)
        logging.error(f"[FAILED] {src_table_original}: {error_msg}")
        return f"[FAILED ] {src_table_original}: {error_msg}"
    finally:
        if local_path and os.path.exists(local_path):
            try:
                time.sleep(1) 
                os.remove(local_path)
            except: pass

# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
if __name__ == "__main__":
    spark = SparkSession.builder \
        .appName("Metadata_Driven_ETL_Fixed") \
        .config("spark.driver.memory", "6g") \
        .config("spark.jars.packages", "com.crealytics:spark-excel_2.12:0.13.7") \
        .getOrCreate()
    spark.sparkContext.setLogLevel("FATAL")

    logging.info("ETL JOB STARTED")
    metadata = load_metadata()
    job_results = []

    if not metadata:
        logging.error("No metadata found in stg.etl_metadata")
        job_results.append("No active metadata found.")
    else:
        for row in metadata:
            try:
                raw_freq = row.get('LOAD_FREQUENCY', 0)
                load_freq = int(raw_freq) if raw_freq is not None else 0
            except:
                load_freq = 0
            
            if load_freq == 0:
                source_table = row.get('SOURCE_TABLE_NAME', 'unknown')
                dest_schema = row.get('DESTINATION_SCHEMA_NAME', 'unknown')
                dest_table = row.get('DESTINATION_DATABASE') or row.get('DESTINATION', 'unknown')
                logging.info(f"Skipping inactive table (LOAD_FREQUENCY=0): {source_table} -> {dest_schema}.{dest_table}")
                log_etl_tracking(source_table, dest_schema, dest_table, "N/A", datetime.now(), "SKIPPED", records=0, error="Load frequency is 0")
                job_results.append(f"[SKIPPED] {source_table}: Load frequency is 0")
                continue

            res = process_metadata_row(row)
            job_results.append(res)
       
    # Email
    final_report = "\n".join(job_results)
    status_subj = "SUCCESS" if "[FAILED ]" not in final_report else "COMPLETED WITH ERRORS"
    send_etl_summary_email(f"ETL Job Report: {status_subj}", final_report)

    spark.stop()