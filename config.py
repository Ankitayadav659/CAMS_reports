"""Central configuration — every tunable number/name lives here."""

APP_NAME = 'CAMS Commission Automation System'

# ───────────── Step 1 : column names (CAMS comm file headers) ─────────────
DATE_COL     = 'TRADE_DATE'
BROK_COL     = 'BRKAGE_RAT'
AMC_COL      = 'AMC_CODE'
SCHEME_COL   = 'SCHEME_COD'
TXN_TYPE_COL = 'TRXN_TYPE'
INV_COL      = 'INV_NAME'
FOLIO_COL    = 'FOLIO_NO'

KEEP_COLS = [DATE_COL, AMC_COL, SCHEME_COL, FOLIO_COL, TXN_TYPE_COL, BROK_COL, INV_COL]
UNIQUE_ID_COLS = [AMC_COL, SCHEME_COL, FOLIO_COL, INV_COL, TXN_TYPE_COL]

# ───────────── Step 1 : thresholds (UNCHANGED from v5.0) ─────────────
THRESHOLD      = -0.2    # minimum drop (negative) to be a candidate
RECOVERY_RATIO = 0.80    # future quarter must recover 80 % of the drop
MIN_TXN        = 2       # ignore quarters with fewer valid transactions

# ───────────── Step 1 : date parsing ─────────────
# v5.0 converted every cell to text and then called to_datetime(dayfirst=True).  For REAL Excel date
# cells that swaps day and month whenever day <= 12 (5-Jan-2024 became 1-May-2024) and can silently
# drop rows whose day is > 12.  The default (False) keeps real dates as real dates, parses ISO text
# year-first and other text day-first.  Set True ONLY to reproduce the old v5.0 behaviour exactly.
LEGACY_V5_DATE_PARSING = False

# ───────────── Excel limits / streaming ─────────────
EXCEL_MAX_ROWS   = 1_048_576
STEP1_HEADER_ROWS = 2                      # legend row + header row
STEP1_ROWS_PER_SHEET = EXCEL_MAX_ROWS - STEP1_HEADER_ROWS
RAW_SHEET_NAME   = 'Raw Data (Highlighted)'
REPORT_SHEET_NAME = 'Discrepancy Report'
WRITE_CHUNK      = 100_000

# ───────────── Step 2 ─────────────
INDEX_SHEET_NAME = 'Investor Sheet Index'
FILE_INDEX_SHEET_NAME = 'File Index'
EXCLUDED_LOG_SHEET_NAME = 'Excluded Rows Log'
BROKERAGE_COL = 'Brokerage'                # new column inserted after BRKAGE_RAT
NEG_FILL = '#FFCCCC'                       # light red for negative Brokerage rows
STEP2_ROWS_PER_SHEET = EXCEL_MAX_ROWS - 1  # data rows per investor sheet (1 header row)

# Output layout of Step 2
#   True  -> ONE SEPARATE EXCEL FILE PER INVESTOR (inside the folder INVESTOR_FOLDER_NAME next to the index file)
#   False -> everything in one workbook (Investor_Wise_Final.xlsx, split into _Part2.. only when too large)
SEPARATE_FILES_PER_INVESTOR = True
INVESTOR_FOLDER_NAME = 'Investor_Excels'

# Extra explanation features (Step 2)
ADD_RATE_CHANGE_COLUMN = True      # new column "Rate Change" right after Brokerage  (▲ Increased / ▼ Decreased / Same)
RATE_CHANGE_COL = 'Rate Change'
ADD_RATE_SUMMARY = True            # "Rate Summary" sheet (periods table + charts) per client; separate-file mode only

# Output workbook splitting (deterministic: groups are packed in sorted order)
MAX_SHEETS_PER_WORKBOOK = 200              # includes index sheets; keeps Windows file handles safe
MAX_CELLS_PER_WORKBOOK  = 30_000_000       # rows x columns budget per workbook

# Source Excel reader: 'auto' | 'openpyxl' | 'calamine'
#   openpyxl = true streaming, lowest RAM (slower)
#   calamine = much faster, but keeps the sheet in native memory
#   auto     = calamine when the file is smaller than CALAMINE_MAX_FILE_MB, else openpyxl
READER = 'auto'
CALAMINE_MAX_FILE_MB = 60

# Investor/AMC validation when several records share SCHEME+FOLIO+TXN
STRICT_INVESTOR_VALIDATION = True

SQLITE_BATCH = 5000
