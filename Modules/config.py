# ============================================
# Jump Host & Credentials
# ============================================
JUMP_HOST = "10.112.250.6"
CRED_TARGET = "MyApp/ADM"

# ============================================
# Connection & Retry Behavior
# ============================================
CONNECTION_TIMEOUT = 25        # SSH connection timeout (seconds)
BANNER_TIMEOUT = 30            # Banner timeout (seconds)
READ_TIMEOUT = 60              # Command read timeout (seconds)
RETRY_ATTEMPTS = 3             # Max connection retry attempts
RETRY_BASE_WAIT = 1            # Base wait time for exponential backoff (seconds)

# ============================================
# TextFSM & Parsing
# ============================================
# Path to NET_TEXTFSM templates (if None, uses environment variable NET_TEXTFSM)
TEXTFSM_PATH = None
USE_TEXTFSM = True             # Enable TextFSM parsing for show interfaces

# ============================================
# Port Classification & Stale Detection
# ============================================
DEFAULT_STALE_DAYS = 30        # Default threshold for stale access port detection
ENABLE_PORT_CATEGORIZATION = True  # Enable port category classification

# ============================================
# Export & Output
# ============================================
DEFAULT_OUTPUT_FORMAT = "xlsx" # Default export format (xlsx, csv, json)
MIN_EXCEL_COLUMN_WIDTH = 8     # Minimum Excel column width
MAX_EXCEL_COLUMN_WIDTH = 60    # Maximum Excel column width

# ============================================
# Concurrency
# ============================================
DEFAULT_WORKERS = 10           # Default max concurrent device sessions

# ============================================
# Alert Thresholds
# ============================================
ERROR_COUNTER_THRESHOLD = 100  # Flag as critical if input/output errors exceed this