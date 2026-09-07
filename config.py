"""
config.py
Central configuration for the certificate/email application.
All secrets (SMTP credentials) are read from environment variables /
a .env file — never hard-code passwords here.
"""

import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Load .env explicitly from this file's own directory — do NOT rely on
# python-dotenv's automatic search, which depends on the current working
# directory and can silently find nothing if the app is launched from a
# different folder.
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
ENV_FILE = BASE_DIR / ".env"

if not ENV_FILE.exists():
    print(
        f"WARNING: no .env file found at {ENV_FILE}\n"
        f"         SMTP credentials and other settings will fall back to defaults.\n"
        f"         Create a .env file at that exact path (copy .env.example -> .env).",
        file=sys.stderr,
    )

load_dotenv(ENV_FILE, override=True)

EXCEL_PATH = Path(os.getenv("EXCEL_PATH", BASE_DIR / "data" / "participants.xlsx"))
PPTX_TEMPLATE_PATH = Path(os.getenv("PPTX_TEMPLATE_PATH", BASE_DIR / "data" / "certificate_template.pptx"))

EMAIL_WITH_TIME_TXT = Path(os.getenv("EMAIL_WITH_TIME_TXT", BASE_DIR / "data" / "email_with_time.txt"))
EMAIL_WITHOUT_TIME_TXT = Path(os.getenv("EMAIL_WITHOUT_TIME_TXT", BASE_DIR / "data" / "email_without_time.txt"))

OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", BASE_DIR / "output"))
PPTX_WORK_DIR = OUTPUT_DIR / "pptx_work"   # per-person filled-in .pptx (intermediate)
PDF_DIR = OUTPUT_DIR / "pdfs"              # final per-person certificate PDFs
LOG_CSV = OUTPUT_DIR / "run_log.csv"       # success/failure log for each row

# ---------------------------------------------------------------------------
# Excel column names (must match the header row exactly)
# ---------------------------------------------------------------------------
COL_NAME = "Name"
COL_DISTANCE = "Distance"
COL_TIME = "Time"
COL_DATE = "Date"
COL_EMAIL = "email_id"

REQUIRED_COLUMNS = [COL_NAME, COL_DISTANCE, COL_TIME, COL_DATE, COL_EMAIL]

# ---------------------------------------------------------------------------
# Placeholder tokens expected inside the PPTX template and the .txt bodies.
# Using {{double_braces}} avoids accidentally matching a stray occurrence of
# a plain word like "Time" inside other certificate text.
# ---------------------------------------------------------------------------
PLACEHOLDERS = {
    "{{Name}}": COL_NAME,
    "{{Distance}}": COL_DISTANCE,
    "{{Time}}": COL_TIME,
    "{{Date}}": COL_DATE,
}

# ---------------------------------------------------------------------------
# SMTP / email settings — set these via environment variables or a .env file:
#   SMTP_HOST=smtp.gmail.com
#   SMTP_PORT=587
#   SMTP_USERNAME=you@example.com
#   SMTP_PASSWORD=app-specific-password
#   SMTP_FROM_NAME=Race Organizers
#   EMAIL_SUBJECT_WITH_TIME=Your Race Certificate
#   EMAIL_SUBJECT_WITHOUT_TIME=Your Race Update
# ---------------------------------------------------------------------------
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USERNAME = os.getenv("SMTP_USERNAME", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SMTP_FROM_NAME = os.getenv("SMTP_FROM_NAME", "Race Organizers")

EMAIL_SUBJECT_WITH_TIME = os.getenv("EMAIL_SUBJECT_WITH_TIME", "Your Race Certificate")
EMAIL_SUBJECT_WITHOUT_TIME = os.getenv("EMAIL_SUBJECT_WITHOUT_TIME", "Your Race Update")

# Path to the LibreOffice binary used to convert filled .pptx -> .pdf,
# used only when CONVERSION_METHOD is "libreoffice" (or "auto" falls back
# to it on non-Windows machines / when PowerPoint COM isn't available).
# On Windows this is usually "soffice.exe"; on Mac it may be
# "/Applications/LibreOffice.app/Contents/MacOS/soffice".
SOFFICE_BIN = os.getenv("SOFFICE_BIN", "soffice")

# How to convert the filled .pptx into a .pdf:
#   "auto"        -> use PowerPoint COM automation on Windows if available,
#                    otherwise fall back to LibreOffice
#   "powerpoint"  -> force PowerPoint COM automation (Windows + PowerPoint required)
#   "libreoffice" -> force headless LibreOffice conversion
CONVERSION_METHOD = os.getenv("CONVERSION_METHOD", "auto")

# Dry run: if True, emails are composed and logged but NOT actually sent.
# Useful for testing the pipeline before going live.
DRY_RUN = os.getenv("DRY_RUN", "true").lower() == "true"