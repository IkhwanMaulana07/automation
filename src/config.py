import os
from pathlib import Path
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # Built-in lightweight .env loader fallback if python-dotenv is not in Jupyter kernel
    env_file = Path(__file__).resolve().parent.parent / ".env"
    if env_file.exists():
        try:
            with open(env_file, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        os.environ.setdefault(k.strip(), v.strip())
        except Exception:
            pass

BASE_DIR = Path(__file__).resolve().parent.parent

# File Paths
GL_FILE_PATH = os.getenv("GL_FILE_PATH", str(BASE_DIR / "GL - Advances Other - April 2026.xls"))
WP_FILE_PATH = os.getenv("WP_FILE_PATH", str(BASE_DIR / "Working Paper Advances and Prepayment-Soal.xlsx"))
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_EXCEL_PATH = OUTPUT_DIR / "Working_Paper_Result.xlsx"
EXECUTIVE_SUMMARY_PATH = OUTPUT_DIR / "Executive_Summary.md"
RECONCILIATION_REPORT_JSON = OUTPUT_DIR / "reconciliation_report.json"

# API & Credentials
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME", "Advance Settlement Reconciliation - SouthCity")

# Model configuration
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")
