"""
FTP(S) -> PostgreSQL pipeline
Streams files directly from FTP into Postgres via COPY FROM STDIN.
No local disk writes. Supports single file, multiple files, or pattern match.
Previews first 5 lines of every file before loading, and logs every
run (per file) into a tracking table.
"""

import io
import logging
from datetime import datetime
from ftplib import FTP_TLS
import os 
import psycopg2
import dotenv
dotenv.load_dotenv()

CONFIG = {
    "ftp_host": os.getenv("FTP_HOST"),
    "ftp_user": os.getenv( "FTP_USER"),
    "ftp_password": os.getenv("FTP_PASSWORD"),
    "ftp_default_dir": "/DP",

    "pg_host":os.getenv("PG_HOST"), 
    "pg_port": 5432,
    "pg_db": os.getenv("PG_DB"),
    "pg_user":os.getenv("PG_USER") ,
    "pg_password": os.getenv("PG_PASSWORD"),
}

# ---- adjust these to your actual load ----
TARGET_SCHEMA = "sales"
TARGET_TABLE_NAME = "sales_billing_daily"
TARGET_TABLE = f"{TARGET_SCHEMA}.{TARGET_TABLE_NAME}"

TRACK_SCHEMA = "sales"
TRACK_TABLE_NAME = "load_log"
TRACK_TABLE = f"{TRACK_SCHEMA}.{TRACK_TABLE_NAME}"

TABLE_COLUMNS = [
    "billing_document",
    "billing_type",
    "sales_organization",
    "distribution_channel",
    "division",
    "billing_date",
    "creation_date",
    "sold_to_customer",
    "ship_to_customer",
    "material",
    "billing_quantity_sku",
    "base_unit_of_measure",
    "cw_quantity",
    "cw_uom",
    "net_amount",
    "plant",
    "sloc",
    "salesman_id",
    "load_current_timestamp",
]

FILE_PATTERN = ".dat"     # used only when TARGET_FILENAMES is empty/None
FILE_DELIMITER = "~"
HAS_HEADER = False
PREVIEW_LINES = 5

# --- file selection ---
# Set one or more filenames to load only those. Leave as [] to load
# every file on the FTP path matching FILE_PATTERN.
TARGET_FILENAMES = ["SALES_H2_2025.dat","SALES_H1_2025.dat"]   # e.g. ["a.dat", "b.dat"] or []

FILE_COLUMNS = [c for c in TABLE_COLUMNS if c != "load_current_timestamp"]
# -------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
log = logging.getLogger("ftp_to_pg")


def connect_ftp() -> FTP_TLS:
    ftp = FTP_TLS()
    ftp.connect(CONFIG["ftp_host"], 21, timeout=30)
    ftp.login(CONFIG["ftp_user"], CONFIG["ftp_password"])
    ftp.prot_p()
    if CONFIG.get("ftp_default_dir"):
        ftp.cwd(CONFIG["ftp_default_dir"])
    log.info("Connected to FTP: %s (%s)", CONFIG["ftp_host"], CONFIG["ftp_default_dir"])
    return ftp


def connect_pg():
    return psycopg2.connect(
        host=CONFIG["pg_host"],
        port=CONFIG["pg_port"],
        dbname=CONFIG["pg_db"],
        user=CONFIG["pg_user"],
        password=CONFIG["pg_password"],
    )


def ensure_target_table(conn):
    with conn.cursor() as cur:
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {TARGET_SCHEMA};")
        cur.execute(f"""
            CREATE TABLE IF NOT EXISTS {TARGET_TABLE} (
                billing_document        TEXT,
                billing_type            TEXT,
                sales_organization      TEXT,
                distribution_channel    TEXT,
                division                TEXT,
                billing_date            TEXT,
                creation_date           TEXT,
                sold_to_customer        TEXT,
                ship_to_customer        TEXT,
                material                TEXT,
                billing_quantity_sku    TEXT,
                base_unit_of_measure    TEXT,
                cw_quantity             TEXT,
                cw_uom                  TEXT,
                net_amount              TEXT,
                plant                   TEXT,
                sloc                    TEXT,
                salesman_id             TEXT,
                load_current_timestamp  TIMESTAMP DEFAULT NOW()
            );
        """)
    conn.commit()


def ensure_track_table(conn):
    """Tracking table: one row per file per run, with counts + status."""
    with conn.cursor() as cur:
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {TRACK_SCHEMA};")
        cur.execute(f"""
            CREATE TABLE IF NOT EXISTS {TRACK_TABLE} (
                id               SERIAL PRIMARY KEY,
                run_id           TIMESTAMP,
                file_name        TEXT,
                target_table     TEXT,
                row_count        INTEGER,
                status           TEXT,          -- SUCCESS / FAILED / SKIPPED
                error_message    TEXT,
                started_at       TIMESTAMP,
                finished_at      TIMESTAMP,
                duration_seconds NUMERIC
            );
        """)
    conn.commit()


def log_file_result(conn, run_id, file_name, row_count, status,
                     error_message, started_at, finished_at):
    duration = (finished_at - started_at).total_seconds()
    with conn.cursor() as cur:
        cur.execute(f"""
            INSERT INTO {TRACK_TABLE}
                (run_id, file_name, target_table, row_count, status,
                 error_message, started_at, finished_at, duration_seconds)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (run_id, file_name, TARGET_TABLE, row_count, status,
              error_message, started_at, finished_at, duration))
    conn.commit()


def list_files(ftp: FTP_TLS) -> list:
    """
    Returns files to load:
    - if TARGET_FILENAMES is non-empty, only those (existence-checked)
    - else every file on FTP matching FILE_PATTERN
    """
    all_files = ftp.nlst()

    if TARGET_FILENAMES:
        found, missing = [], []
        for fn in TARGET_FILENAMES:
            (found if fn in all_files else missing).append(fn)
        if missing:
            log.error("Not found on FTP, skipping: %s", missing)
        log.info("Targeting %d file(s): %s", len(found), found)
        return found

    candidates = [f for f in all_files if f.lower().endswith(FILE_PATTERN)]
    log.info("Found %d matching files on FTP", len(candidates))
    return candidates


def fetch_file_to_memory(ftp: FTP_TLS, remote_name: str) -> io.BytesIO:
    buf = io.BytesIO()
    ftp.retrbinary(f"RETR {remote_name}", buf.write)
    buf.seek(0)
    log.info("Fetched %s into memory (%d bytes)", remote_name, buf.getbuffer().nbytes)
    return buf


def preview_file(remote_name: str, lines: list):
    log.info("--- preview: %s (first %d lines) ---", remote_name, PREVIEW_LINES)
    for line in lines[:PREVIEW_LINES]:
        print(line)
    log.info("--- end preview: %s ---", remote_name)


def load_buffer_to_postgres(conn, buf: io.BytesIO, remote_name: str) -> int:
    text = buf.read().decode("utf-8")
    lines = text.splitlines()
    if not lines:
        log.warning("%s is empty, skipping", remote_name)
        return 0

    if HAS_HEADER:
        columns = [c.strip() for c in lines[0].split(FILE_DELIMITER)]
        data_lines = lines[1:]
    else:
        columns = FILE_COLUMNS
        data_lines = lines

    preview_file(remote_name, data_lines)

    col_clause = f"({', '.join(columns)})" if columns else ""
    copy_sql = (
        f"COPY {TARGET_TABLE} {col_clause} "
        f"FROM STDIN WITH (FORMAT csv, DELIMITER '{FILE_DELIMITER}')"
    )

    data_stream = io.StringIO("\n".join(data_lines))
    with conn.cursor() as cur:
        cur.copy_expert(copy_sql, data_stream)

    return len(data_lines)


def fetch_all_files(ftp: FTP_TLS, files: list) -> dict:
    """
    Downloads every file into memory up front, then the FTP control
    connection is closed. Keeps it from sitting idle (and timing out /
    dropping with SSLEOFError) while Postgres COPY runs in between fetches.
    Returns {filename: BytesIO}; files that fail to download are skipped
    here and logged, not included in the returned dict.
    """
    buffers = {}
    for remote_name in files:
        try:
            buffers[remote_name] = fetch_file_to_memory(ftp, remote_name)
        except Exception as e:
            log.error("Failed to fetch %s from FTP: %s", remote_name, e)
    return buffers


def run(target_filenames=None):
    global TARGET_FILENAMES
    if target_filenames:
        TARGET_FILENAMES = target_filenames

    run_id = datetime.now()
    conn = connect_pg()
    ensure_target_table(conn)
    ensure_track_table(conn)

    # --- phase 1: FTP - fetch everything into memory, then disconnect ---
    ftp = connect_ftp()
    try:
        files = list_files(ftp)
        buffers = fetch_all_files(ftp, files)
    finally:
        ftp.quit()

    # --- phase 2: Postgres - load each buffer, FTP connection is closed ---
    total_files = total_rows = total_success = total_failed = 0
    try:
        for remote_name, buf in buffers.items():
            started_at = datetime.now()
            total_files += 1
            try:
                with conn:  # transaction per file
                    row_count = load_buffer_to_postgres(conn, buf, remote_name)
                finished_at = datetime.now()
                log.info("Loaded %s (%d rows)", remote_name, row_count)
                log_file_result(conn, run_id, remote_name, row_count,
                                 "SUCCESS", None, started_at, finished_at)
                total_rows += row_count
                total_success += 1
            except Exception as e:
                conn.rollback()
                finished_at = datetime.now()
                log.error("Failed loading %s: %s", remote_name, e)
                log_file_result(conn, run_id, remote_name, 0,
                                 "FAILED", str(e), started_at, finished_at)
                total_failed += 1

        skipped = len(files) - len(buffers)
        if skipped:
            log.warning("%d file(s) skipped (FTP fetch failure)", skipped)
    finally:
        log.info(
            "Run %s complete | files=%d success=%d failed=%d total_rows=%d",
            run_id, total_files, total_success, total_failed, total_rows,
        )
        conn.close()


if __name__ == "__main__":
    import sys
    # pass one or more filenames as CLI args to override TARGET_FILENAMES
    # e.g. python ftp_to_postgres.py SALES_1JAN.dat SALES_2JAN.dat
    cli_files = sys.argv[1:] if len(sys.argv) > 1 else None
    run(cli_files)