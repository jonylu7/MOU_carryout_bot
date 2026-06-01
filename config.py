"""
carry_out_automation/config.py
攜出自動化流程 — 設定檔（所有可調整參數集中在這裡）

個人資訊從 .env 載入；Gemini API 金鑰從 .env.gemini 載入。
"""

import os
from pathlib import Path

from dotenv import load_dotenv

_HERE = Path(__file__).parent
load_dotenv(_HERE / ".env")
load_dotenv(_HERE / ".env.gemini")


# ─────────────────────────────────────────────────
# 路徑設定
# ─────────────────────────────────────────────────
BASE_DIR = Path(os.environ["BASE_DIR"])

TEMPLATE_DIR = _HERE / "template"

# 子題 → 空白範本對應
TEMPLATES = {
    "子題一":       "範例_MOU攜出子題一資料申請YYYYMMDD.docx",
    "子題一子議題":  "範例_MOU攜出子題一子議題資料申請YYYYMMDD.docx",
    "子題二":       "範例_MOU攜出子題二資料申請YYYYMMDD.docx",
    "稅收估計":     "範例_MOU攜出稅收估計資料申請YYYYMMDD.docx",
}
DEFAULT_TEMPLATE = "範例_MOU攜出稅收估計資料申請YYYYMMDD.docx"

CARRYOUT_FOLDER   = BASE_DIR / "攜出專區" / "資料攜出" / "攜出資料夾(統一放這)"
CARRYOUT_FORM_DIR = BASE_DIR / "攜出專區" / "資料攜出" / "攜出單"


# ─────────────────────────────────────────────────
# Google Drive 設定
# ─────────────────────────────────────────────────
GOOGLE_SERVICE_ACCOUNT_KEY = _HERE / os.environ["GOOGLE_SERVICE_ACCOUNT_KEY"]
DRIVE_FILE_ID  = os.environ["DRIVE_FILE_ID"]
SHEET_NAME     = "攜出檔案填寫"


# ─────────────────────────────────────────────────
# Gemini API 設定
# ─────────────────────────────────────────────────
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL   = "gemini-2.0-flash"


# ─────────────────────────────────────────────────
# Email 設定
# ─────────────────────────────────────────────────
EMAIL_SMTP_HOST   = "smtp.gmail.com"
EMAIL_SMTP_PORT   = 587
EMAIL_SENDER      = os.environ["EMAIL_SENDER"]
EMAIL_PASSWORD    = os.environ["EMAIL_PASSWORD"]
EMAIL_COORDINATOR = os.environ["EMAIL_COORDINATOR"]
EMAIL_MOF_TO      = os.environ["EMAIL_MOF_TO"]
_cc_raw           = os.environ.get("EMAIL_MOF_CC", "")
EMAIL_MOF_CC: list[str] = [a.strip() for a in _cc_raw.split(",") if a.strip()]


# ─────────────────────────────────────────────────
# HITL 審核伺服器設定
# ─────────────────────────────────────────────────
HITL_HOST         = "127.0.0.1"
HITL_PORT         = 8765
HITL_SECRET_TOKEN = os.environ.get("HITL_TOKEN", "changeme-please")
HITL_TIMEOUT_SECS = 3600


# ─────────────────────────────────────────────────
# 電子簽章
# ─────────────────────────────────────────────────
SIGNATURE_IMAGE_PATH = _HERE / "signature.png"
SIGNATURE_WIDTH_CM   = 3.0


# ─────────────────────────────────────────────────
# 流程邏輯
# ─────────────────────────────────────────────────
DAYS_AHEAD_FILTER = 14


# ─────────────────────────────────────────────────
# 財資中心子題承辦人
# ─────────────────────────────────────────────────
MOU_COORDINATORS: dict[str, dict] = {
    "子題一":      {"name": "薛秀英", "email": "n107217@fia.gov.tw"},
    "子題二":      {"name": "許巧君", "email": "n108266@fia.gov.tw"},
    "子題一子議題": {"name": "陳靜",   "email": "n108362@fia.gov.tw"},
    "稅收估計":    {"name": "許宏韜",  "email": "N108237@fia.gov.tw"},
}


# ─────────────────────────────────────────────────
# 機構固定資訊
# ─────────────────────────────────────────────────
INSTITUTE_NAME    = "國立政治大學台灣研究中心"
DEFAULT_APPLICANT = os.environ.get("DEFAULT_APPLICANT", "盧威任")
ROC_YEAR_OFFSET   = 1911
