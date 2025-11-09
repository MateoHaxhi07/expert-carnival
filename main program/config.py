# ============================================================================
# RESTAURANT REPORT CONFIGURATION FILE
# ============================================================================

# ============================================================================
# BUSINESS HOURS
# ============================================================================
LUNCH_START_HOUR = 1       # e.g., 11
LUNCH_END_HOUR   = 17      # e.g., 15
DINNER_START_HOUR = 17     # e.g., 18
DINNER_END_HOUR   = 24     # e.g., 22

# ============================================================================
# CURRENCY & FORMATTING
# ============================================================================
CURRENCY_SYMBOL = 'ALL'
THOUSANDS_SEPARATOR = ','   # (kept for future use if you want localized text outputs)
DECIMAL_PLACES = 0          # 0 for integers, 2 for cents

# ============================================================================
# BUSINESS TARGETS/GOALS
# ============================================================================
DAILY_REVENUE_TARGET = 70000
WEEKLY_REVENUE_TARGET = 500000
MONTHLY_REVENUE_TARGET = 2000000
ENABLE_TARGET_TRACKING = True

# ============================================================================
# REPORT CUSTOMIZATION
# ============================================================================
TOP_ITEMS_COUNT = 10
BOTTOM_ITEMS_COUNT = 25
INCLUDE_PREVIOUS_PERIOD = True

# ============================================================================
# GOOGLE SHEETS SETTINGS
# ============================================================================
CREDENTIALS_FILE = 'credentials.json'
SHEET_INDEX = 0

# ============================================================================
# EXCEL OUTPUT SETTINGS
# ============================================================================
REPORT_NAME_PREFIX = 'CEO_Report'
INCLUDE_DATE_IN_FILENAME = True

# Optional classic toggles (kept for compatibility)
INCLUDE_EXECUTIVE_SUMMARY = True
INCLUDE_DAILY_PERFORMANCE = True
INCLUDE_MENU_ANALYSIS = True
INCLUDE_TIME_ANALYSIS = True
INCLUDE_HOURLY_BREAKDOWN = True
INCLUDE_CUSTOMER_BEHAVIOR = True

# NEW sheets (you can toggle these)
INCLUDE_ANALYTIC_SALES_SHEET = True        # full row-level sheet sorted by Datetime ASC
INCLUDE_ALL_ITEMS_BY_QTY_SHEET = True      # single “All Items (Qty Desc)” sheet

# Also write a short .txt summary beside the Excel
CREATE_TEXT_SUMMARY = True

# ============================================================================
# ALERTS & NOTIFICATIONS
# ============================================================================
ALERT_IF_REVENUE_DROPS_BY = 20
ALERT_IF_ORDERS_DROP_BY = 15
SHOW_ALERTS = True

# ============================================================================
# COLUMN MAPPING
# ============================================================================
COLUMN_MAPPING = {
    'Order_ID': 'Order_ID',
    'Article_Name': 'Article_Name',
    'Category': 'Category',
    'Quantity': 'Quantity',
    'Total_Article_Price': 'Total_Article_Price',
    'Datetime': 'Datetime',
    'Seller': 'Seller',
    'Seller Category': 'Seller_Category',  # Sheet col → standard col
    'Buyer_NIPT': 'Buyer_NIPT',
    'Buyer_Name': 'Buyer_Name'
}







INCLUDE_ITEMS_COMPARISON_SHEET = True
MIN_ROWS_FOR_COMPARISON = 1
INCLUDE_DAILY_COMPARISON_SHEET = True
INCLUDE_TIMEPERIOD_COMPARISON_SHEET = True
INCLUDE_HOURLY_COMPARISON_SHEET = True


# ============================================================================
# ADVANCED SETTINGS
# ============================================================================
MINIMUM_ORDERS_PER_ITEM = 1
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'  # fallback parsing will still try if format differs

EXCLUDED_ITEMS = [
    'DISCOUNT',
    'SHERBIM KAMARIERI',
    'SHERBIM DISTANCE'
]

# config.py
EXCLUDE_SELLERS_FROM_MAIN = False   # keep special sellers in main report too

# ============================================================================
# ITEMS COMPARISON SHEET (NEW)
# ============================================================================
INCLUDE_ITEMS_COMPARISON_SHEET = True   # writes "All Items — Comparison (Prev Period)"
MIN_ROWS_FOR_COMPARISON = 1             # require at least this many rows in both periods





# ============================================================================
# SELLER ANALYSIS (NEW!)
# ============================================================================
# Sellers to analyze separately (e.g., catering, wholesale, special accounts)
# These sellers will be:
#   1. EXCLUDED from main report metrics
#   2. GET their own dedicated analysis sheet in the Excel file

SELLERS_TO_ANALYZE_SEPARATELY = []

# Enable separate seller analysis sheets?
INCLUDE_SELLER_ANALYSIS_SHEETS = False

# What details to show in seller sheets
SELLER_SHEET_SHOW_ORDER_IDS = True          # Show individual order IDs
SELLER_SHEET_SHOW_DAILY_BREAKDOWN = True    # Daily performance
SELLER_SHEET_SHOW_ITEM_BREAKDOWN = True     # Items sold breakdown
SELLER_SHEET_SHOW_TIME_ANALYSIS = True      # Lunch/dinner breakdown
SELLER_SHEET_SHOW_BUYER_INFO = True         # Buyer NIPT and name

# ============================================================================
# END OF CONFIGURATION
# ============================================================================




