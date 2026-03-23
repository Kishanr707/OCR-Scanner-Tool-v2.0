# config.py
# ─── Gemini API Configuration ─────────────────────────────────────────────────

GEMINI_API_KEY = "-----------------"        # paste your key here before building

GEMINI_MODEL   = "gemini-2.5-flash"         # model to use for all calls

# ─── Request hardening ────────────────────────────────────────────────────────

REQUEST_TIMEOUT        = 30      # seconds before a request is considered dead
MAX_RETRIES            = 2       # number of times to retry on failure
RETRY_DELAY            = 2       # seconds to wait between retries

# ─── Document handling ────────────────────────────────────────────────────────

DOCX_TEXT_LIMIT        = 800     # max characters from DOCX to send to Gemini
PDF_TEXT_LIMIT         = 1000    # max characters from text-based PDF to send
GARBAGE_THRESHOLD      = 20      # min printable chars to consider text valid

# ─── Excel ────────────────────────────────────────────────────────────────────

EXCEL_FILENAME         = "contacts.xlsx"

# ─── App ──────────────────────────────────────────────────────────────────────

APP_TITLE              = "Visiting Card Scanner"
APP_VERSION            = "2.0.0"
MAX_FILES              = 10
