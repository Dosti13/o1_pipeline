import os
import re
import csv
import fnmatch
import ftplib
import logging
import tempfile
import datetime as dt
from typing import Dict, List

import psycopg2
from pyspark.sql import SparkSession, DataFrame
from pyspark.sql import functions as F
from pyspark.sql.functions import col, lit
import os 
import dotenv 
dotenv.load_dotenv()
# =========================================================
# CONFIG (qnie_new_test)
# =========================================================
CONFIG = {
    "ftp_host": os.getenv("FTP_HOST"),
    "ftp_user": os.getenv("FTP_USER"),
    "ftp_password": os.getenv("FTP_PASSWORD"),
    "ftp_default_dir": "DP",

    "pg_host": os.getenv("PG_HOST"),
    "pg_port": 5432,
    "pg_db": os.getenv("PG_DB"),
    "pg_user": os.getenv("PG_USER"),
    "pg_password": os.getenv("PG_PASSWORD"),

    "jdbc_url": f"jdbc:postgresql://localhost:5432/{os.getenv("PG_DB")}",
    "jdbc_driver": "org.postgresql.Driver",

    "metadata_path": r"C:\Users\HP\Desktop\meta_data_2_extended1.txt",
    "metadata_sep": "\t",
    # pipeline start timestamp (UTC not required if your old log used local; keep consistent)
    "pipeline_start_ts": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
}

# =========================================================
# LOGGING
# =========================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("etl_job_qnie_new_test.log", mode="a", encoding="utf-8")
    ]
)

# =========================================================
# SPARK
# =========================================================
spark = (
    SparkSession.builder
    .appName("MetaDrivenETL_qnie_new_test")
    .master("local[*]")
    .config("spark.driver.memory", "6g")
    .config("spark.executor.memory", "8g")
    .config(
        "spark.jars.packages",
        "org.postgresql:postgresql:42.7.3,"
        "com.crealytics:spark-excel_2.12:0.13.7"
    )
    .getOrCreate()
)
spark.sparkContext.setLogLevel("ERROR")

# =========================================================
# JDBC SETTINGS (for meta.file_log + loads)
# =========================================================
JDBC_URL = CONFIG["jdbc_url"]
JDBC_PROPS = {
    "user": CONFIG["pg_user"],
    "password": CONFIG["pg_password"],
    "driver": CONFIG["jdbc_driver"],
    "stringtype": "unspecified",
}
META_FILE_LOG_TABLE = "meta.file_log"

# =========================================================
# DB HELPERS
# =========================================================
def pg_connect():
    return psycopg2.connect(
        host=CONFIG["pg_host"],
        port=CONFIG["pg_port"],
        dbname=CONFIG["pg_db"],
        user=CONFIG["pg_user"],
        password=CONFIG["pg_password"],
    )

def exec_sql(sql_text: str):
    if not sql_text:
        return

    sql_text = str(sql_text).strip()

    if sql_text.upper() in ("FALSE", "TRUE", "NULL", ""):
        return
    conn = pg_connect()
    conn.autocommit = True
    try:
        with conn.cursor() as cur:
            cur.execute(sql_text)
        logging.info(f"Executed SQL: {sql_text}")
    finally:
        conn.close()

def get_table_columns(schema: str, table: str) -> List[str]:
    conn = pg_connect()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT column_name
                FROM information_schema.columns
                WHERE table_schema=%s AND table_name=%s
                ORDER BY ordinal_position
            """, (schema, table))
            return [r[0] for r in cur.fetchall()]
    finally:
        conn.close()

def align_schema(df: DataFrame, schema: str, table: str) -> DataFrame:
    target_cols = get_table_columns(schema, table)
    if not target_cols:
        return df

    df_map = {c.lower(): c for c in df.columns}

    for t in target_cols:
        tl = t.lower()
        if tl not in df_map:
            df = df.withColumn(t, lit(None).cast("string"))
        else:
            actual = df_map[tl]
            if actual != t:
                df = df.withColumnRenamed(actual, t)

    return df.select(*target_cols)

# =========================================================
# meta.file_log WRITER (Spark JDBC)
# =========================================================
def log_file_event_spark(
    pipeline_start_ts_utc: str,
    source_system: str,
    feed_name: str,
    file_name: str,
    remote_path: str,
    local_path: str,
    status: str,
    rows_loaded: int = 0,
    message: str = ""
):
    """
    Inserts 1 row into meta.file_log.
    Must NOT break pipeline if logging fails.
    """
    try:
        msg = "" if message is None else str(message)
        if len(msg) > 8000:
            msg = msg[:8000]

        df = spark.createDataFrame(
            [(source_system, feed_name, file_name, remote_path, local_path,
              status, int(rows_loaded or 0), msg, pipeline_start_ts_utc)],
            ["source_system", "feed_name", "file_name", "remote_path", "local_path",
             "status", "rows_loaded", "message", "pipeline_start_ts_utc"]
        ) \
        .withColumn("run_ts", F.current_timestamp()) \
        .withColumn("pipeline_start_ts_utc", F.to_timestamp("pipeline_start_ts_utc")) \
        .select(
            "run_ts", "pipeline_start_ts_utc", "source_system", "feed_name", "file_name",
            "remote_path", "local_path", "status", "rows_loaded", "message"
        )

        df.write.mode("append").jdbc(JDBC_URL, META_FILE_LOG_TABLE, properties=JDBC_PROPS)

    except Exception as e:
        logging.warning(f"[FILE_LOG] insert failed: {e}")

# =========================================================
# FTP HELPERS
# =========================================================
def ftp_connect():
    # NOTE: You are using ftplib.FTP (not FTP_TLS). Keep same as your current working setup.
    ftp = ftplib.FTP(CONFIG["ftp_host"])
    ftp.login(CONFIG["ftp_user"], CONFIG["ftp_password"])
    return ftp

def normalize_remote_dir(rd: str) -> str:
    rd = (rd or "").strip()
    if not rd:
        rd = CONFIG["ftp_default_dir"]
    rd = rd.replace("\\", "/")
    if rd.startswith("/"):
        rd = rd[1:]
    if rd.endswith("/"):
        rd = rd[:-1]
    return rd

def ftp_list(remote_dir: str) -> List[str]:
    ftp = ftp_connect()
    try:
        rd = normalize_remote_dir(remote_dir)
        if rd:
            ftp.cwd(rd)
        return ftp.nlst()
    finally:
        try:
            ftp.quit()
        except Exception:
            pass

def ftp_download(remote_dir: str, filename: str) -> str:
    suffix = os.path.splitext(filename)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp_path = tmp.name
    tmp.close()

    ftp = ftp_connect()
    try:
        rd = normalize_remote_dir(remote_dir)
        if rd:
            ftp.cwd(rd)
        with open(tmp_path, "wb") as f:
            ftp.retrbinary("RETR " + filename, f.write)
        logging.info(f"Downloaded FTP {rd}/{filename} -> {tmp_path}")
        return tmp_path
    finally:
        try:
            ftp.quit()
        except Exception:
            pass

# =========================================================
# METADATA LOADER
# =========================================================
def load_metadata(path: str, sep: str) -> List[Dict]:
    rows = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=sep)
        for r in reader:
            rr = {k: (v.strip() if isinstance(v, str) else v) for k, v in r.items()}
            try:
                if int(rr.get("LOAD_FREQUENCY") or 0) == 0:
                    continue
            except Exception:
                continue
            rows.append(rr)

    def pr(x):
        try:
            return int(x.get("PRIORITY") or 9999)
        except Exception:
            return 9999

    rows.sort(key=pr)
    return rows

# =========================================================
# EXCEL REFRESH (WMS)
# =========================================================
# def refresh_excel_if_needed(local_path: str):
#     try:
#         import time
#         import pythoncom
#         import win32com.client as win32
#     except Exception:
#         logging.info("[WMS] win32com not available - skipping refresh.")
#         return

#     if not local_path or not os.path.exists(local_path):
#         logging.warning(f"[WMS] File not found, cannot refresh: {local_path}")
#         return

#     logging.info(f"[WMS] Refreshing Excel: {local_path}")
#     pythoncom.CoInitialize()
#     excel = win32.DispatchEx("Excel.Application")
#     excel.Visible = False
#     excel.DisplayAlerts = False
#     excel.AskToUpdateLinks = False
#     try:
#         wb = excel.Workbooks.Open(local_path, UpdateLinks=0)
#         wb.RefreshAll()
#         while True:
#             refreshing = getattr(wb, "Refreshing", False)
#             calc = getattr(excel, "CalculateState", 0)
#             if (not refreshing) and calc == 0:
#                 break
#             time.sleep(1)
#         wb.Save()
#         wb.Close(SaveChanges=False)
#         logging.info("[WMS] Excel refresh done.")
#     finally:
#         excel.Quit()
#         pythoncom.CoUninitialize()

# =========================================================
# TRANSFORMS
# =========================================================
def _parse_month_to_yyyy_mm_01(col_expr: F.Column) -> F.Column:
    s = F.trim(col_expr)
    digits = F.regexp_replace(s, r"\D+", "")
    digits6 = F.when(F.length(digits) >= 6, F.substring(digits, 1, 6)).otherwise(lit(""))
    yyyy = F.substring(digits6, 1, 4)
    mm = F.substring(digits6, 5, 2)
    mm_int = mm.cast("int")
    valid = (F.length(digits6) == 6) & (mm_int >= 1) & (mm_int <= 12)
    ym = F.concat_ws("-", yyyy, F.lpad(mm, 2, "0"), F.lit("01"))
    return F.when(valid, F.to_date(ym, "yyyy-MM-dd")).otherwise(F.lit(None).cast("date"))

def _clean_numeric(col_expr: F.Column) -> F.Column:
    s = F.trim(col_expr)
    s = F.when(s.isNull() | (s == "") | (F.upper(s) == "NULL"), lit(None)).otherwise(s)
    s = F.regexp_replace(s, r"^\((.*)\)$", r"-\1")
    s = F.regexp_replace(s, r"[,\s]", "")
    s = F.regexp_replace(s, r"^([+-]?[0-9]*\.?[0-9]+)-$", r"-\1")
    return s

def _po_fix(df: DataFrame) -> DataFrame:
    candidates = ["po", "line", "vendor", "plant", "material"]
    for c in candidates:
        if c in df.columns:
            x = F.trim(F.col(c).cast("string"))
            x = F.regexp_replace(x, r"[,\s]", "")
            x = F.regexp_replace(x, r"\.0+$", "")
            df = df.withColumn(c, x)
    return df

def _clean_date(col_expr: F.Column) -> F.Column:
    s = F.trim(col_expr.cast("string"))
    s = F.when(s.isNull() | (s == "") | (F.upper(s) == "NULL"), F.lit(None)).otherwise(s)

    ts = F.coalesce(
        F.to_timestamp(s, "M/d/yy H:mm"),
        F.to_timestamp(s, "M/d/yyyy H:mm"),
        F.to_timestamp(s, "M/d/yy HH:mm"),
        F.to_timestamp(s, "M/d/yyyy HH:mm"),
        F.to_timestamp(s, "M/d/yy H:mm:ss"),
        F.to_timestamp(s, "M/d/yyyy H:mm:ss"),
        F.to_timestamp(s, "yyyy-MM-dd HH:mm:ss"),
        F.to_timestamp(s, "yyyy-MM-dd H:mm:ss"),
        F.to_timestamp(s, "yyyy-MM-dd HH:mm"),
        F.to_timestamp(s, "yyyy-MM-dd H:mm"),
    )

    d = F.coalesce(
        F.to_date(s, "M/d/yy"),
        F.to_date(s, "M/d/yyyy"),
        F.to_date(s, "yyyy-MM-dd"),
        F.to_date(s, "yyyy/MM/dd"),
        F.to_date(s, "yyyy.MM.dd"),
    )

    return F.when(ts.isNotNull(), F.date_format(ts, "yyyy-MM-dd HH:mm:ss")) \
            .when(d.isNotNull(), F.date_format(d, "yyyy-MM-dd")) \
            .otherwise(F.lit(None).cast("string"))

def apply_transforms(df: DataFrame, meta: Dict) -> DataFrame:
    t = (meta.get("TRANSFORMS") or "").strip()
    if not t:
        return df

    steps = [s.strip() for s in t.split(";") if s.strip()]

    for step in steps:
        if step == "trim_all":
            for c in df.columns:
                df = df.withColumn(c, F.trim(F.col(c)))

        elif step == "po_fix":
            df = _po_fix(df)

        elif step.startswith("date_parse:"):
            arg = step.split(":", 1)[1].strip()
            cols = [c.strip() for c in arg.split("|") if c.strip()]
            for c in cols:
                if c in df.columns:
                    df = df.withColumn(c, _clean_date(F.col(c)))
                else:
                    logging.warning(f"[TRANSFORM] date_parse: column not found: {c}")

        elif step.startswith("month_parse:"):
            arg = step.split(":", 1)[1].strip()
            src_dst = arg.split("->")
            src_col = src_dst[0].strip()
            dst_col = src_dst[1].strip() if len(src_dst) > 1 else src_col

            if src_col in df.columns:
                df = df.withColumn(dst_col, _parse_month_to_yyyy_mm_01(F.col(src_col)))
            else:
                logging.warning(f"[TRANSFORM] month_parse: column not found: {src_col}")

        elif step.startswith("numeric_clean:"):
            arg = step.split(":", 1)[1].strip()
            cols = [c.strip() for c in arg.split("|") if c.strip()]
            for c in cols:
                if c in df.columns:
                    df = df.withColumn(c, _clean_numeric(F.col(c)))
                else:
                    logging.warning(f"[TRANSFORM] numeric_clean: column not found: {c}")

        elif step == "drop_invalid_months":
            if "per_year" in df.columns:
                df = df.filter(F.col("per_year").isNotNull())
            else:
                logging.warning("[TRANSFORM] drop_invalid_months: per_year not found")

        else:
            logging.warning(f"[TRANSFORM] Unknown step ignored: {step}")

    return df

# =========================================================
# READ FILE (handles headerless + multi delimiter)
# =========================================================
def should_treat_as_headerless(meta: Dict, filename: str) -> bool:
    reader_type = (meta.get("READER_TYPE") or "").lower()
    if reader_type == "excel":
        return False

    hdr = (meta.get("HEADER") or "").strip()
    ext = os.path.splitext(filename)[1].lower()

    if hdr and ext in [".dat", ".txt"]:
        return True
    if hdr:
        return True
    return False

def read_source_file(local_path: str, meta: Dict, filename: str) -> DataFrame:
    reader_type = (meta.get("READER_TYPE") or "csv").lower()
    delimiter = meta.get("FILE_DELIMITER") or "~"
    multi_delim = meta.get("MULTI_DELIMITER") or ""

    headerless = should_treat_as_headerless(meta, filename)
    header_option = "false" if headerless else "true"

    if reader_type == "excel":
        sheet = meta.get("SHEET_NAME") or "Sheet1"
        df = (
            spark.read.format("com.crealytics.spark.excel")
            .option("header", "true")
            .option("dataAddress", f"'{sheet}'!A1")
            .option("inferSchema", "false")
            .load(local_path)
        )
        return df

    if multi_delim.strip():
        raw = spark.read.text(local_path)
        parts = F.split(raw["value"], re.escape(multi_delim.strip()))
        hdr = (meta.get("HEADER") or "").strip()
        n = len([h.strip() for h in hdr.split(",") if h.strip()]) if hdr else 200
        cols = [parts.getItem(i).alias(f"_c{i}") for i in range(n)]
        return raw.select(*cols)

    return (
        spark.read
        .option("header", header_option)
        .option("delimiter", delimiter)
        .option("inferSchema", "false")
        .option("encoding", "utf-8")
        .csv(local_path)
    )

def apply_headers_if_provided(df: DataFrame, meta: Dict) -> DataFrame:
    hdr = (meta.get("HEADER") or "").strip()
    if not hdr:
        return df
    headers = [h.strip() for h in hdr.split(",") if h.strip()]
    if not headers:
        return df

    if len(df.columns) < len(headers):
        for i in range(len(df.columns), len(headers)):
            df = df.withColumn(f"_pad_{i}", lit(None).cast("string"))

    return df.select(df.columns[:len(headers)]).toDF(*headers)

# =========================================================
# CORE PROCESSOR (with meta.file_log)
# =========================================================
def process_row(meta: Dict):
    source_system = (meta.get("TABLE_GROUP") or "UNKNOWN").strip()
    feed_name = (meta.get("SOURCE_TABLE_NAME") or "").strip()

    src = meta.get("SOURCE_TABLE_NAME")
    mode = (meta.get("LOAD_TABLE") or "").strip().upper()

    dest_schema = (meta.get("DESTINATION_SCHEMA") or "").strip()
    dest_table = (meta.get("DESTINATION_TABLE") or "").strip()

    logging.info(f"=== Processing {src} -> {dest_schema}.{dest_table} ===")

    # For logs
    is_local = str(meta.get("IS_LOCAL_FILE") or "").upper() == "TRUE"
    local_path = ""
    filename = ""
    remote_path_for_log = ""
    remote_dir = meta.get("REMOTE_DIR") or CONFIG["ftp_default_dir"]

    # POST_SQL_ONLY
    if mode in ("POST_SQL_ONLY","FUNC","PROC"):
        sql_to_run = (meta.get("SQL_TEXT") or meta.get("POST_SQL") or "").strip()

        try:
            exec_sql(meta.get(sql_to_run))
            log_file_event_spark(
                CONFIG["pipeline_start_ts"], source_system, feed_name,

                src or feed_name, "DATABASE", "", "SUCCESS", 0,
                "POST_SQL_ONLY executed"
            )
        except Exception as e:
            log_file_event_spark(
                CONFIG["pipeline_start_ts"], source_system, feed_name,
                src or feed_name, "DATABASE", "", "FAILED", 0, repr(e)
            )
            raise
        return

    exec_sql(meta.get("PRE_SQL", ""))

    # Determine file
    if is_local:
        local_path = meta.get("LOCAL_PATH") or ""
        if not local_path:
            raise Exception("LOCAL_PATH missing for local file row.")
        # if (meta.get("READER_TYPE") or "").lower() == "excel":
        #       refresh_excel_if_needed(local_path)
        filename = os.path.basename(local_path)
        remote_path_for_log = "LOCAL_DESKTOP"
    else:
        files = ftp_list(remote_dir)

        patt = src
        if "*" in patt or "?" in patt:
            matched = sorted([f for f in files if fnmatch.fnmatch(f, patt)], key=str.lower)
        else:
            matched = [f for f in files if f.lower() == patt.lower()]

        if not matched:
            raise Exception(f"No file found in FTP dir={remote_dir} for pattern={patt}")

        latest_only = str(meta.get("FTP_LATEST_ONLY") or "").upper() == "TRUE"
        if latest_only:
            rx = meta.get("LATEST_REGEX") or ""
            if rx:
                rxx = re.compile(rx)
                cand = [f for f in matched if rxx.match(f)]
                matched = cand or matched

        pick = matched[-1]
        local_path = ftp_download(remote_dir, pick)
        filename = pick
        remote_path_for_log = f"/{normalize_remote_dir(remote_dir)}/{filename}"

    # STARTED log
    log_file_event_spark(
        pipeline_start_ts_utc=CONFIG["pipeline_start_ts"],
        source_system=source_system,
        feed_name=feed_name,
        file_name=filename or (src or feed_name),
        remote_path=remote_path_for_log,
        local_path=local_path,
        status="STARTED",
        rows_loaded=0,
        message=f"dest={dest_schema}.{dest_table} mode={mode}"
    )

    try:
        df = read_source_file(local_path, meta, filename)
        df = apply_headers_if_provided(df, meta)

        for c in df.columns:
            df = df.withColumn(c, col(c).cast("string"))

        df = apply_transforms(df, meta)

        audit_col = (meta.get("AUDIT_TS_COL") or "").strip()
        if audit_col:
            df = df.withColumn(audit_col, lit(CONFIG["pipeline_start_ts"]))

        if dest_schema and dest_table:
            df = align_schema(df, dest_schema, dest_table)

        rows_loaded = df.count()

        jdbc_opts = {
            "url": CONFIG["jdbc_url"],
            "dbtable": f'"{dest_schema}"."{dest_table}"',
            "user": CONFIG["pg_user"],
            "password": CONFIG["pg_password"],
            "driver": CONFIG["jdbc_driver"],
            "stringtype": "unspecified",
        }

        df.write.mode("append").format("jdbc").options(**jdbc_opts).save()

        exec_sql(meta.get("POST_SQL", ""))
        logging.info(f"Loaded {src} into {dest_schema}.{dest_table}")

        log_file_event_spark(
            CONFIG["pipeline_start_ts"], source_system, feed_name,
            filename or (src or feed_name),
            remote_path_for_log, local_path, "SUCCESS", rows_loaded, ""
        )

    except Exception as e:
        log_file_event_spark(
            CONFIG["pipeline_start_ts"], source_system, feed_name,
            filename or (src or feed_name),
            remote_path_for_log, local_path, "FAILED", 0, repr(e)
        )
        raise

    finally:
        if (not is_local) and local_path and os.path.exists(local_path):
            try:
                os.remove(local_path)
            except Exception:
                pass

# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":
    logging.info("JOB STARTED (qnie_new_test)")
    logging.info(f"PIPELINE_START_TS={CONFIG['pipeline_start_ts']}")

    meta_rows = load_metadata(CONFIG["metadata_path"], sep=CONFIG["metadata_sep"])
    if not meta_rows:
        raise SystemExit("No active metadata rows found (LOAD_FREQUENCY=1).")

    for r in meta_rows:
        try:
            process_row(r)
        except Exception as e:
            logging.error(f"FAILED {r.get('SOURCE_TABLE_NAME')}: {e}", exc_info=True)

    logging.info("JOB FINISHED")