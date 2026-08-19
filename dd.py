from typing import Dict, List
import ftplib
import dotenv 
import os 
dotenv.load_dotenv()

CONFIG = {
    "ftp_host": os.getenv("FTP_HOST"),
    "ftp_user": os.getenv("FTP_USER"),
    "ftp_password": os.getenv("FTP_PASSWORD"),
    "ftp_default_dir": "DP",
}
print(CONFIG["ftp_host"])
# def ftp_connect():



#     # NOTE: You are using ftplib.FTP (not FTP_TLS). Keep same as your current working setup.
#     ftp = ftplib.FTP(CONFIG["ftp_host"])
#     ftp.login(CONFIG["ftp_user"], CONFIG["ftp_password"])
#     return ftp

# def normalize_remote_dir(rd: str) -> str:
#     rd = (rd or "").strip()
#     if not rd:
#         rd = CONFIG["ftp_default_dir"]
#     rd = rd.replace("\\", "/")
#     if rd.startswith("/"):
#         rd = rd[1:]
#     if rd.endswith("/"):
#         rd = rd[:-1]
#     return rd

# def ftp_list(remote_dir: str) -> List[str]:
#     ftp = ftp_connect()
#     try:
#         rd = normalize_remote_dir(remote_dir)
#         if rd:
#             print(rd,"rd")
#             ftp.cwd(rd)
#         files = ftp.nlst()
#         print(ftp.dir())
#         return files

#     finally:
#         try:
#             ftp.quit()
#         except Exception:
#             pass
# remote_dir =  CONFIG["ftp_default_dir"]

# ftp_list(remote_dir)        