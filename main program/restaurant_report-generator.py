"""
RESTAURANT WEEKLY CEO REPORT GENERATOR (config-driven)
======================================================
- Reads Google Sheets data via Service Account
- Uses config.py for ALL tunables (hours, currency, targets, sheet, etc.)
- Cleans data, filters by Seller Category and date range, computes KPIs, and exports Excel (+ optional TXT)
- Separate analysis for specific sellers (e.g., catering, wholesale)
- "Analytic Sales" sheet and "All Items (Qty Desc)" sheet
- NEW: Fixed Seller Category menu with explicit 'NaN' option
- NEW: All Items — Comparison (Prev Period) sheet
"""

import warnings
warnings.filterwarnings('ignore')

import pandas as pd
from collections import Counter
from itertools import combinations
from datetime import datetime, timedelta

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ===== import ALL settings from config.py =====
from config import (
    # Business hours / currency / formatting
    LUNCH_START_HOUR, LUNCH_END_HOUR, DINNER_START_HOUR, DINNER_END_HOUR,
    CURRENCY_SYMBOL, THOUSANDS_SEPARATOR, DECIMAL_PLACES,

    # Targets and switches
    ENABLE_TARGET_TRACKING, DAILY_REVENUE_TARGET, WEEKLY_REVENUE_TARGET, MONTHLY_REVENUE_TARGET,

    # Report composition
    TOP_ITEMS_COUNT, BOTTOM_ITEMS_COUNT, INCLUDE_PREVIOUS_PERIOD,
    REPORT_NAME_PREFIX, INCLUDE_DATE_IN_FILENAME, CREATE_TEXT_SUMMARY,

    # Google Sheets
    CREDENTIALS_FILE, SHEET_INDEX,

    # Alerts
    ALERT_IF_REVENUE_DROPS_BY, ALERT_IF_ORDERS_DROP_BY, SHOW_ALERTS,

    # Columns & parsing
    COLUMN_MAPPING, MINIMUM_ORDERS_PER_ITEM, DATE_FORMAT, EXCLUDED_ITEMS,

    # Seller analysis
    SELLERS_TO_ANALYZE_SEPARATELY, INCLUDE_SELLER_ANALYSIS_SHEETS,
    SELLER_SHEET_SHOW_ORDER_IDS, SELLER_SHEET_SHOW_DAILY_BREAKDOWN,
    SELLER_SHEET_SHOW_ITEM_BREAKDOWN, SELLER_SHEET_SHOW_TIME_ANALYSIS,
    SELLER_SHEET_SHOW_BUYER_INFO,

    # NEW sheet toggles
    INCLUDE_ANALYTIC_SALES_SHEET, INCLUDE_ALL_ITEMS_BY_QTY_SHEET,

    # Main inclusion flag for special sellers
    EXCLUDE_SELLERS_FROM_MAIN,
)

# ======== NEW (add these two in config.py if missing) ========
try:
    from config import INCLUDE_ITEMS_COMPARISON_SHEET, MIN_ROWS_FOR_COMPARISON
except ImportError:
    # sensible defaults if not present in your config.py yet
    INCLUDE_ITEMS_COMPARISON_SHEET = True
    MIN_ROWS_FOR_COMPARISON = 1
# =============================================================


# -------------------------------------------------------------------------
# 1) Google Sheets connection
# -------------------------------------------------------------------------
def connect_to_google_sheets(spreadsheet_url: str) -> pd.DataFrame:
    """Authorize with service account and load the target worksheet into a DataFrame."""
    print("📊 Connecting to Google Sheets...")

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_url(spreadsheet_url)
        ws = sheet.get_worksheet(SHEET_INDEX)

        records = ws.get_all_records()
        df = pd.DataFrame(records)

        # normalize/mapping column names
        if COLUMN_MAPPING:
            df = df.rename(columns=COLUMN_MAPPING)

        # drop excluded items (discounts, service, etc.)
        if EXCLUDED_ITEMS and "Article_Name" in df.columns:
            df = df[~df["Article_Name"].isin(EXCLUDED_ITEMS)]

        print(f"✅ Successfully loaded {len(df)} records from worksheet '{ws.title}'")
        return df

    except Exception as e:
        print(f"❌ Error connecting to Google Sheets: {e}")
        print("   • Check spreadsheet URL")
        print("   • Share the Sheet with your service account email")
        print(f"   • Ensure {CREDENTIALS_FILE} exists next to this script")
        raise

# -------------------------------------------------------------------------
# 2) Cleaning & enrichment
# -------------------------------------------------------------------------
def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Coerce types, parse datetime, derive helpers, ensure minimal schema."""
    print("\n🧹 Cleaning data...")

    required = {"Order_ID", "Article_Name", "Quantity", "Total_Article_Price", "Datetime"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    df = df.copy()

    # Parse datetime (allow configured format, tolerate bad rows)
    try:
        df["Datetime"] = pd.to_datetime(df["Datetime"], errors="coerce", format=DATE_FORMAT)
    except Exception:
        df["Datetime"] = pd.to_datetime(df["Datetime"], errors="coerce")

    # Remove bad rows
    df = df.dropna(subset=["Datetime"])

    # Numerics
    for col in ["Quantity", "Total_Article_Price"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Derive helpers
    df["Date"] = df["Datetime"].dt.date
    df["Hour"] = df["Datetime"].dt.hour
    df["DayOfWeek"] = df["Datetime"].dt.day_name()
    df["WeekNumber"] = df["Datetime"].dt.isocalendar().week
    df["Month"] = df["Datetime"].dt.month
    df["Year"] = df["Datetime"].dt.year

    # Time buckets
    def bucket(hour: int) -> str:
        if LUNCH_START_HOUR <= hour < LUNCH_END_HOUR:
            return "Lunch"
        if DINNER_START_HOUR <= hour < DINNER_END_HOUR:
            return "Dinner"
        return "Other"

    df["TimePeriod"] = df["Hour"].apply(bucket)

    # Keep Seller_Category as-is (do NOT normalize). We need real NaN support.
    if "Seller_Category" in df.columns:
        df["Seller_Category"] = df["Seller_Category"].where(~df["Seller_Category"].isna(), other=pd.NA)

    # Optional: filter rare items (by number of orders)
    if MINIMUM_ORDERS_PER_ITEM > 1:
        order_counts = df.groupby("Article_Name")["Order_ID"].nunique()
        keep_items = order_counts[order_counts >= MINIMUM_ORDERS_PER_ITEM].index
        df = df[df["Article_Name"].isin(keep_items)]

    print(f"✅ Data cleaned. {len(df)} valid rows remain.")
    return df

# -------------------------------------------------------------------------
# 2a) FIXED Seller Category menu with explicit NaN
# -------------------------------------------------------------------------
def choose_seller_categories_fixed(df: pd.DataFrame):
    """
    Fixed menu:
      1) All
      2) Bar
      3) Delivery
      4) Restaurant
      5) NaN (blank/missing or literal 'NaN')
      6) Bar+Restaurant+Delivery (quick combo)
    Returns (selection_keys, label)
    """
    print("\n👥 SELECT SELLER CATEGORY")
    print("=" * 50)
    print("1) All")
    print("2) Bar")
    print("3) Delivery")
    print("4) Restaurant")
    print("5) NaN (blank)")
    print("6) Bar + Restaurant + Delivery")
    print("=" * 50)
    raw = input("Enter a number or comma-list (e.g. 3 or 2,4,5): ").strip()

    if not raw:
        raw = "1"

    try:
        idxs = [int(x) for x in raw.split(",") if x.strip() != ""]
    except ValueError:
        print("⚠️  Invalid input. Using All.")
        idxs = [1]

    index_to_key = {
        1: "ALL",
        2: "BAR",
        3: "DELIVERY",
        4: "RESTAURANT",
        5: "NAN",
        6: "BRD",  # NEW: Bar+Restaurant+Delivery
    }
    keys = [index_to_key[i] for i in idxs if i in index_to_key]

    # If All included, ignore other picks
    if "ALL" in keys or not keys:
        print("✅ Category filter: All")
        return ["ALL"], "All"

    labels_map = {
        "BAR": "Bar",
        "DELIVERY": "Delivery",
        "RESTAURANT": "Restaurant",
        "NAN": "NaN",
        "BRD": "Bar+Restaurant+Delivery",
    }
    label = "+".join([labels_map[k] for k in keys])
    print(f"✅ Category filter: {label}")
    return keys, label

def filter_by_seller_category_fixed(df: pd.DataFrame, keys) -> pd.DataFrame:
    """
    Apply fixed-key filtering:
      - BAR         -> 'Bar'
      - DELIVERY    -> 'Delivery'
      - RESTAURANT  -> 'Restaurant'
      - NAN         -> isna() OR '' OR literal 'nan'
      - BRD         -> Bar OR Restaurant OR Delivery   (NEW)
      - ALL         -> no filter
    """
    if "ALL" in keys or "Seller_Category" not in df.columns:
        return df

    sc = df["Seller_Category"]

    # normalize string comparisons without touching original df
    def norm_series(s: pd.Series) -> pd.Series:
        s_str = s.astype("string")
        return s_str.str.strip().str.lower()

    s_norm = norm_series(sc)
    mask_total = pd.Series(False, index=df.index)

    # atomic categories
    mask_bar        = (s_norm == "bar")
    mask_delivery   = (s_norm == "delivery")
    mask_restaurant = (s_norm == "restaurant")
    mask_nan_true   = sc.isna()
    mask_empty      = (s_norm == "")
    mask_text_nan   = (s_norm == "nan")

    # compose from keys
    if "BAR" in keys:
        mask_total |= mask_bar
    if "DELIVERY" in keys:
        mask_total |= mask_delivery
    if "RESTAURANT" in keys:
        mask_total |= mask_restaurant
    if "NAN" in keys:
        mask_total |= (mask_nan_true | mask_empty | mask_text_nan)
    if "BRD" in keys:
        mask_total |= (mask_bar | mask_delivery | mask_restaurant)

    return df[mask_total].copy()

# -------------------------------------------------------------------------
# 2b) Separate seller data (respects EXCLUDE_SELLERS_FROM_MAIN)
# -------------------------------------------------------------------------
def separate_seller_data(df: pd.DataFrame):
    """
    Split data into:
    1. Main report data (depending on EXCLUDE_SELLERS_FROM_MAIN)
    2. Dictionary of seller-specific DataFrames
    Returns: (df_main, dict_of_seller_dfs)
    """
    if not SELLERS_TO_ANALYZE_SEPARATELY or not INCLUDE_SELLER_ANALYSIS_SHEETS:
        return df, {}

    print("\n🔍 Separating seller data...")

    if "Seller" not in df.columns:
        print("⚠️  Warning: 'Seller' column not found. Skipping seller separation.")
        return df, {}

    seller_dfs = {}
    for seller in SELLERS_TO_ANALYZE_SEPARATELY:
        seller_df = df[df["Seller"] == seller].copy()
        if len(seller_df) > 0:
            seller_dfs[seller] = seller_df
            print(f"   📊 {seller}: {len(seller_df)} rows")
        else:
            print(f"   ⚠️  {seller}: No data found")

    if EXCLUDE_SELLERS_FROM_MAIN:
        df_main = df[~df["Seller"].isin(SELLERS_TO_ANALYZE_SEPARATELY)].copy()
        print(f"✅ Main report: {len(df_main)} rows (excluded {len(df) - len(df_main)} seller rows)")
    else:
        df_main = df.copy()
        print(f"✅ Main report: {len(df_main)} rows (kept all sellers)")

    return df_main, seller_dfs

# -------------------------------------------------------------------------
# 3) Date range helpers
# -------------------------------------------------------------------------
def get_date_range():
    """Console menu for picking a date window."""
    print("\n📅 SELECT DATE RANGE FOR REPORT")
    print("=" * 50)
    print("1) This Week (Mon-Sun)")
    print("2) Last Week")
    print("3) This Month")
    print("4) Last Month")
    print("5) Last 7 Days")
    print("6) Last 30 Days")
    print("7) Custom (YYYY-MM-DD .. YYYY-MM-DD)")
    print("=" * 50)

    choice = input("Enter choice (1-7): ").strip()
    today = datetime.now().date()

    if choice == "1":
        start = today - timedelta(days=today.weekday())
        end = start + timedelta(days=6)
        label = "This Week"
    elif choice == "2":
        this_mon = today - timedelta(days=today.weekday())
        start = this_mon - timedelta(days=7)
        end = start + timedelta(days=6)
        label = "Last Week"
    elif choice == "3":
        start = today.replace(day=1)
        next_month = (start.replace(day=28) + timedelta(days=4)).replace(day=1)
        end = next_month - timedelta(days=1)
        label = "This Month"
    elif choice == "4":
        first_this = today.replace(day=1)
        end = first_this - timedelta(days=1)
        start = end.replace(day=1)
        label = "Last Month"
    elif choice == "5":
        end = today
        start = today - timedelta(days=6)
        label = "Last 7 Days"
    elif choice == "6":
        end = today
        start = today - timedelta(days=29)
        label = "Last 30 Days"
    elif choice == "7":
        start = datetime.strptime(input("Start (YYYY-MM-DD): ").strip(), "%Y-%m-%d").date()
        end = datetime.strptime(input("End   (YYYY-MM-DD): ").strip(), "%Y-%m-%d").date()
        label = f"Custom ({start} to {end})"
    else:
        end = today
        start = today - timedelta(days=6)
        label = "Last 7 Days (default)"

    print(f"\n✅ {label}: {start} → {end}")
    return start, end, label

def filter_by_date(df: pd.DataFrame, start_date, end_date) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"]).dt.date
    mask = (df["Date"] >= start_date) & (df["Date"] <= end_date)
    return df[mask]

# -------------------------------------------------------------------------
# 4) Metrics engine
# -------------------------------------------------------------------------
class ReportMetrics:
    def __init__(self, df_current: pd.DataFrame, df_previous: pd.DataFrame | None):
        self.df_current = df_current
        self.df_previous = df_previous
        self.metrics = {}

    def calculate_executive(self):
        cur = self.df_current
        total_revenue = float(cur["Total_Article_Price"].sum())
        total_orders = int(cur["Order_ID"].nunique())
        avg_order_value = total_revenue / total_orders if total_orders else 0.0

        num_days = max(1, len(pd.unique(cur["Date"])))
        orders_per_day = total_orders / num_days

        result = {
            "total_revenue": total_revenue,
            "total_orders": total_orders,
            "avg_order_value": avg_order_value,
            "orders_per_day": orders_per_day,
            "num_days": num_days,
        }

        if self.df_previous is not None and len(self.df_previous):
            prev_rev = float(self.df_previous["Total_Article_Price"].sum())
            prev_ord = int(self.df_previous["Order_ID"].nunique())
            result["revenue_change"] = ((total_revenue - prev_rev) / prev_rev * 100) if prev_rev else 0.0
            result["orders_change"] = ((total_orders - prev_ord) / prev_ord * 100) if prev_ord else 0.0

            if SHOW_ALERTS:
                result["revenue_drop_alert"] = (result["revenue_change"] < -abs(ALERT_IF_REVENUE_DROPS_BY))
                result["orders_drop_alert"] = (result["orders_change"] < -abs(ALERT_IF_ORDERS_DROP_BY))

        self.metrics["executive"] = result
        return result

    def calculate_daily(self):
        cur = self.df_current
        if cur.empty:
            self.metrics["daily"] = {"data": pd.DataFrame(), "best_day": None, "worst_day": None}
            return self.metrics["daily"]

        daily = (
            cur.groupby(["Date", "DayOfWeek"])
            .agg(Revenue=("Total_Article_Price", "sum"),
                 Orders=("Order_ID", "nunique"))
            .reset_index()
            .sort_values("Date")
        )
        daily["AvgOrderValue"] = daily.apply(lambda r: r["Revenue"] / r["Orders"] if r["Orders"] else 0, axis=1)

        best = daily.iloc[daily["Revenue"].idxmax()] if len(daily) else None
        worst = daily.iloc[daily["Revenue"].idxmin()] if len(daily) else None

        self.metrics["daily"] = {"data": daily, "best_day": best, "worst_day": worst}
        return self.metrics["daily"]

    def calculate_menu(self):
        cur = self.df_current
        if cur.empty:
            self.metrics["menu"] = {"all_items": pd.DataFrame(), "top": pd.DataFrame(), "bottom": pd.DataFrame()}
            return self.metrics["menu"]

        menu = (
            cur.groupby("Article_Name")
            .agg(Quantity=("Quantity", "sum"),
                 Revenue=("Total_Article_Price", "sum"),
                 NumOrders=("Order_ID", "nunique"))
            .reset_index()
            .sort_values("Revenue", ascending=False)
        )
        total_rev = float(menu["Revenue"].sum()) or 1.0
        menu["RevenuePercent"] = menu["Revenue"].div(total_rev).fillna(0) * 100

        self.metrics["menu"] = {
            "all_items": menu,
            "top": menu.head(TOP_ITEMS_COUNT),
            "bottom": menu.tail(BOTTOM_ITEMS_COUNT),
        }
        return self.metrics["menu"]

    def calculate_time(self):
        cur = self.df_current
        periods = (
            cur.groupby("TimePeriod")
            .agg(Revenue=("Total_Article_Price", "sum"),
                 Orders=("Order_ID", "nunique"))
            .reset_index()
        )
        periods["AvgOrderValue"] = periods.apply(lambda r: r["Revenue"] / r["Orders"] if r["Orders"] else 0, axis=1)
        total_rev = float(periods["Revenue"].sum()) or 1.0
        periods["RevenuePercent"] = periods["Revenue"].div(total_rev).fillna(0) * 100

        hourly = (
            cur.groupby("Hour")
            .agg(Revenue=("Total_Article_Price", "sum"),
                 Orders=("Order_ID", "nunique"))
            .reset_index()
            .sort_values("Hour")
        )

        self.metrics["time"] = {"periods": periods, "hourly": hourly}
        return self.metrics["time"]

    def calculate_basket(self):
        cur = self.df_current
        pairs = Counter()
        for items in cur.groupby("Order_ID")["Article_Name"].apply(list):
            uniq = sorted(set(items))
            for a, b in combinations(uniq, 2):
                pairs[(a, b)] += 1
        self.metrics["basket"] = {"top_combos": pairs.most_common(10)}
        return self.metrics["basket"]

    def calculate_all(self):
        print("\n" + "=" * 60)
        print("🔢 CALCULATING METRICS")
        print("=" * 60)
        self.calculate_executive()
        self.calculate_daily()
        self.calculate_menu()
        self.calculate_time()
        self.calculate_basket()
        print("✅ Metrics calculated.")
        return self.metrics

# -------------------------------------------------------------------------
# 4a) Items table builders + comparison
# -------------------------------------------------------------------------
def _build_all_items(df: pd.DataFrame) -> pd.DataFrame:
    """
    Returns All Items table with Quantity, Revenue, NumOrders and RevenuePercent.
    """
    if df.empty:
        return pd.DataFrame(columns=["Article_Name","Quantity","Revenue","NumOrders","RevenuePercent"])

    t = (
        df.groupby("Article_Name")
          .agg(Quantity=("Quantity", "sum"),
               Revenue=("Total_Article_Price", "sum"),
               NumOrders=("Order_ID", "nunique"))
          .reset_index()
    )
    total_rev = float(t["Revenue"].sum()) or 1.0
    t["RevenuePercent"] = t["Revenue"].div(total_rev).fillna(0) * 100
    return t


def build_items_comparison(cur_df: pd.DataFrame, prev_df: pd.DataFrame) -> pd.DataFrame:
    """
    Compare All Items of current vs previous period.
    Columns:
      Article_Name, Qty_Current, Rev_Current, Qty_Prev, Rev_Prev,
      ΔQty, ΔRev, RevShare_Current, RevShare_Prev, Δpp
    """
    cur = _build_all_items(cur_df).rename(columns={
        "Quantity":"Qty_Current",
        "Revenue":"Rev_Current",
        "RevenuePercent":"RevShare_Current"
    })[["Article_Name","Qty_Current","Rev_Current","RevShare_Current"]]

    prev = _build_all_items(prev_df).rename(columns={
        "Quantity":"Qty_Prev",
        "Revenue":"Rev_Prev",
        "RevenuePercent":"RevShare_Prev"
    })[["Article_Name","Qty_Prev","Rev_Prev","RevShare_Prev"]]

    out = cur.merge(prev, on="Article_Name", how="outer").fillna(0)
    out["ΔQty"] = out["Qty_Current"] - out["Qty_Prev"]
    out["ΔRev"] = out["Rev_Current"] - out["Rev_Prev"]
    out["Δpp"] = out["RevShare_Current"] - out["RevShare_Prev"]

    # Sort by current revenue, then qty; you can change to sort by ΔRev if preferred
    out = out.sort_values(["Rev_Current","Qty_Current"], ascending=[False, False]).reset_index(drop=True)
    return out

# -------------------------------------------------------------------------
# 4b) Seller-specific analysis
# -------------------------------------------------------------------------
class SellerAnalysis:
    def __init__(self, seller_name: str, df: pd.DataFrame):
        self.seller_name = seller_name
        self.df = df
        self.analysis = {}

    def calculate_all(self):
        print(f"\n📊 Analyzing seller: {self.seller_name}")

        if self.df.empty:
            print(f"   ⚠️  No data for {self.seller_name}")
            return self.analysis

        total_revenue = float(self.df["Total_Article_Price"].sum())
        total_orders = int(self.df["Order_ID"].nunique())
        avg_order = total_revenue / total_orders if total_orders else 0
        total_items = int(self.df["Quantity"].sum())

        self.analysis["summary"] = {
            "total_revenue": total_revenue,
            "total_orders": total_orders,
            "avg_order_value": avg_order,
            "total_items": total_items,
        }

        if SELLER_SHEET_SHOW_ORDER_IDS:
            orders = (
                self.df.groupby("Order_ID")
                .agg(
                    Date=("Date", "first"),
                    Revenue=("Total_Article_Price", "sum"),
                    Items=("Quantity", "sum"),
                    NumArticles=("Article_Name", "nunique")
                )
                .reset_index()
                .sort_values("Date")
            )

            if SELLER_SHEET_SHOW_BUYER_INFO:
                buyer_cols = []
                if "Buyer_Name" in self.df.columns:
                    buyer_info = self.df.groupby("Order_ID")["Buyer_Name"].first()
                    orders = orders.merge(buyer_info.rename("Buyer_Name"), left_on="Order_ID", right_index=True, how="left")
                    buyer_cols.append("Buyer_Name")
                if "Buyer_NIPT" in self.df.columns:
                    nipt_info = self.df.groupby("Order_ID")["Buyer_NIPT"].first()
                    orders = orders.merge(nipt_info.rename("Buyer_NIPT"), left_on="Order_ID", right_index=True, how="left")
                    buyer_cols.append("Buyer_NIPT")
                if buyer_cols:
                    cols = ["Order_ID"] + buyer_cols + [c for c in orders.columns if c not in ["Order_ID"] + buyer_cols]
                    orders = orders[cols]

            self.analysis["orders"] = orders

        if SELLER_SHEET_SHOW_DAILY_BREAKDOWN:
            daily = (
                self.df.groupby(["Date", "DayOfWeek"])
                .agg(
                    Revenue=("Total_Article_Price", "sum"),
                    Orders=("Order_ID", "nunique"),
                    Items=("Quantity", "sum")
                )
                .reset_index()
                .sort_values("Date")
            )
            self.analysis["daily"] = daily

        if SELLER_SHEET_SHOW_ITEM_BREAKDOWN:
            items = (
                self.df.groupby("Article_Name")
                .agg(
                    Quantity=("Quantity", "sum"),
                    Revenue=("Total_Article_Price", "sum"),
                    NumOrders=("Order_ID", "nunique")
                )
                .reset_index()
                .sort_values("Revenue", ascending=False)
            )
            total_rev = float(items["Revenue"].sum()) or 1.0
            items["RevenuePercent"] = items["Revenue"].div(total_rev).fillna(0) * 100
            self.analysis["items"] = items

        if SELLER_SHEET_SHOW_TIME_ANALYSIS:
            time = (
                self.df.groupby("TimePeriod")
                .agg(
                    Revenue=("Total_Article_Price", "sum"),
                    Orders=("Order_ID", "nunique")
                )
                .reset_index()
            )
            self.analysis["time"] = time

        print(f"   ✅ Analysis complete: {total_orders} orders, {total_revenue:,.0f} {CURRENCY_SYMBOL}")
        return self.analysis

# -------------------------------------------------------------------------
# 5) Builders for Analytic Sales
# -------------------------------------------------------------------------
def build_analytic_sales(df: pd.DataFrame) -> pd.DataFrame:
    t = df.copy()
    t["Datetime"] = pd.to_datetime(t["Datetime"], errors="coerce")
    t = t.dropna(subset=["Datetime"])

    preferred_cols = [
        "Date", "Datetime", "Hour", "DayOfWeek", "TimePeriod",
        "Order_ID", "Seller", "Seller_Category",
        "Buyer_Name", "Buyer_NIPT",
        "Article_Name", "Category",
        "Quantity", "Total_Article_Price"
    ]
    cols = [c for c in preferred_cols if c in t.columns]
    t = t[cols].sort_values("Datetime", ascending=True).reset_index(drop=True)
    return t

# -------------------------------------------------------------------------
# 6) Excel / TXT output
# -------------------------------------------------------------------------
def generate_excel_report(
    metrics: dict,
    seller_analyses: dict,
    period_name: str,
    output_filename: str,
    df_current_filtered: pd.DataFrame,
    df_previous_filtered: pd.DataFrame | None
):
    print("\n📄 Generating Excel report...")

    with pd.ExcelWriter(output_filename, engine="xlsxwriter") as writer:
        book = writer.book

        def fmt_currency(ws, col_idxs):
            decimal_part = f".{'0'*DECIMAL_PLACES}" if DECIMAL_PLACES > 0 else ""
            fmt = book.add_format({"num_format": f'#,##0{decimal_part} "{CURRENCY_SYMBOL}"'})
            for idx in col_idxs:
                ws.set_column(idx, idx, 14, fmt)

        def fmt_headers(ws):
            header = book.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1})
            ws.set_row(0, None, header)
            ws.freeze_panes(1, 0)

        # Executive
        ex = metrics["executive"]
        exec_rows = [
            ("Total Revenue", ex["total_revenue"]),
            ("Total Orders", ex["total_orders"]),
            ("Average Order Value", ex["avg_order_value"]),
            ("Orders per Day", round(ex["orders_per_day"], 1)),
        ]
        if "revenue_change" in ex:
            exec_rows.append(("Revenue Change vs Previous", f'{ex.get("revenue_change", 0):.1f}%'))
            exec_rows.append(("Orders Change vs Previous", f'{ex.get("orders_change", 0):.1f}%'))

        exec_df = pd.DataFrame(exec_rows, columns=["Metric","Value"])
        exec_df.to_excel(writer, sheet_name="Executive Summary", index=False)
        ws = writer.sheets["Executive Summary"]
        fmt_headers(ws)

        # Daily Performance
        daily_df = metrics["daily"]["data"]
        if not daily_df.empty:
            daily_df.to_excel(writer, sheet_name="Daily Performance", index=False)
            ws = writer.sheets["Daily Performance"]
            fmt_headers(ws)
            # Revenue, AvgOrderValue columns
            colmap = {name: i for i, name in enumerate(daily_df.columns)}
            rev_idx = colmap.get("Revenue", None)
            aov_idx = colmap.get("AvgOrderValue", None)
            to_fmt = [i for i in [rev_idx, aov_idx] if i is not None]
            if to_fmt:
                fmt_currency(ws, to_fmt)

        # All Items (Qty Desc)
        if "menu" in metrics and not metrics["menu"]["all_items"].empty and INCLUDE_ALL_ITEMS_BY_QTY_SHEET:
            items_all = metrics["menu"]["all_items"].copy()
            if "Quantity" in items_all.columns:
                items_all = items_all.sort_values("Quantity", ascending=False)
            items_all.to_excel(writer, sheet_name="All Items (Qty Desc)", index=False)
            ws = writer.sheets["All Items (Qty Desc)"]
            fmt_headers(ws)
            if "Revenue" in items_all.columns:
                rev_idx = list(items_all.columns).index("Revenue")
                fmt_currency(ws, [rev_idx])

        # NEW: All Items — Comparison (Prev Period)
        if INCLUDE_ITEMS_COMPARISON_SHEET and df_previous_filtered is not None:
            comp_df = build_items_comparison(df_current_filtered, df_previous_filtered)
            if not comp_df.empty and len(df_current_filtered) >= MIN_ROWS_FOR_COMPARISON:
                comp_df.to_excel(writer, sheet_name="All Items — Comparison", index=False)
                ws = writer.sheets["All Items — Comparison"]
                fmt_headers(ws)
                # Currency columns
                for col_name in ["Rev_Current", "Rev_Prev", "ΔRev"]:
                    if col_name in comp_df.columns:
                        fmt_currency(ws, [list(comp_df.columns).index(col_name)])

        # Lunch vs Dinner
        time_df = metrics["time"]["periods"]
        if not time_df.empty:
            time_df.to_excel(writer, sheet_name="Lunch vs Dinner", index=False)
            ws = writer.sheets["Lunch vs Dinner"]
            fmt_headers(ws)
            colmap = {name: i for i, name in enumerate(time_df.columns)}
            to_fmt = [i for i in [colmap.get("Revenue"), colmap.get("AvgOrderValue")] if i is not None]
            if to_fmt:
                fmt_currency(ws, to_fmt)

        # Hourly Analysis
        hourly_df = metrics["time"]["hourly"]
        if not hourly_df.empty:
            hourly_df.to_excel(writer, sheet_name="Hourly Analysis", index=False)
            ws = writer.sheets["Hourly Analysis"]
            fmt_headers(ws)
            if "Revenue" in hourly_df.columns:
                fmt_currency(ws, [list(hourly_df.columns).index("Revenue")])

        # Analytic Sales (row-level)
        if df_current_filtered is not None and not df_current_filtered.empty and INCLUDE_ANALYTIC_SALES_SHEET:
            analytic_df = build_analytic_sales(df_current_filtered)
            if not analytic_df.empty:
                analytic_df.to_excel(writer, sheet_name="Analytic Sales", index=False)
                ws = writer.sheets["Analytic Sales"]
                fmt_headers(ws)
                if "Total_Article_Price" in analytic_df.columns:
                    price_idx = list(analytic_df.columns).index("Total_Article_Price")
                    fmt_currency(ws, [price_idx])

        # Seller-specific sheets
        if seller_analyses and INCLUDE_SELLER_ANALYSIS_SHEETS:
            for seller_name, analysis in seller_analyses.items():
                if not analysis:
                    continue
                safe_name = seller_name[:25]

                summary = analysis.get("summary", {})
                summary_df = pd.DataFrame({
                    "Metric": ["Total Revenue", "Total Orders", "Avg Order Value", "Total Items"],
                    "Value": [
                        summary.get("total_revenue", 0),
                        summary.get("total_orders", 0),
                        summary.get("avg_order_value", 0),
                        summary.get("total_items", 0)
                    ]
                })
                sheet_name = f"{safe_name}-Summary"
                summary_df.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                fmt_headers(ws)

                if "orders" in analysis:
                    orders_df = analysis["orders"]
                    sheet_name = f"{safe_name}-Orders"
                    orders_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]
                    fmt_headers(ws)
                    if "Revenue" in orders_df.columns:
                        rev_idx = list(orders_df.columns).index("Revenue")
                        fmt_currency(ws, [rev_idx])

                if "daily" in analysis:
                    daily_s = analysis["daily"]
                    sheet_name = f"{safe_name}-Daily"
                    daily_s.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]
                    fmt_headers(ws)
                    if "Revenue" in daily_s.columns:
                        rev_idx = list(daily_s.columns).index("Revenue")
                        fmt_currency(ws, [rev_idx])

                if "items" in analysis:
                    items_s = analysis["items"]
                    sheet_name = f"{safe_name}-Items"
                    items_s.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]
                    fmt_headers(ws)
                    if "Revenue" in items_s.columns:
                        rev_idx = list(items_s.columns).index("Revenue")
                        fmt_currency(ws, [rev_idx])

                if "time" in analysis:
                    time_s = analysis["time"]
                    sheet_name = f"{safe_name}-Time"
                    time_s.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]
                    fmt_headers(ws)
                    if "Revenue" in time_s.columns:
                        rev_idx = list(time_s.columns).index("Revenue")
                        fmt_currency(ws, [rev_idx])

    print(f"✅ Report saved: {output_filename}")

def maybe_write_text_summary(metrics: dict, period_name: str, filename_no_ext: str):
    if not CREATE_TEXT_SUMMARY:
        return
    ex = metrics["executive"]
    with open(f"{filename_no_ext}.txt", "w", encoding="utf-8") as f:
        f.write(f"{period_name}\n")
        f.write("-" * 40 + "\n")
        f.write(f"Revenue: {ex['total_revenue']:,.0f} {CURRENCY_SYMBOL}\n")
        f.write(f"Orders : {ex['total_orders']:,}\n")
        aov = ex['avg_order_value'] if ex['total_orders'] else 0
        f.write(f"AOV    : {aov:,.0f} {CURRENCY_SYMBOL}\n")
        if "revenue_change" in ex:
            f.write(f"Rev Δ  : {ex['revenue_change']:.1f}%\n")
            f.write(f"Ord Δ  : {ex['orders_change']:.1f}%\n")
        if ENABLE_TARGET_TRACKING:
            f.write("\nTargets\n")
            f.write(f"- Daily : {DAILY_REVENUE_TARGET:,.0f} {CURRENCY_SYMBOL}\n")
            f.write(f"- Weekly: {WEEKLY_REVENUE_TARGET:,.0f} {CURRENCY_SYMBOL}\n")
            f.write(f"- Monthly:{MONTHLY_REVENUE_TARGET:,.0f} {CURRENCY_SYMBOL}\n")
    print(f"📝 Text summary saved: {filename_no_ext}.txt")

# -------------------------------------------------------------------------
# 7) Main flow
# -------------------------------------------------------------------------
def main():
    print("\n" + "=" * 60)
    print("🍽️  RESTAURANT CEO REPORT GENERATOR (config-driven)")
    print("=" * 60)
    print("This will build a multi-sheet Excel report from your Google Sheet.\n")

    sheet_url = input("Paste your Google Sheet URL: ").strip()

    # load & clean
    df = connect_to_google_sheets(sheet_url)
    df = clean_data(df)

    # Fixed Seller Category selection (includes NaN)
    keys, cats_label = choose_seller_categories_fixed(df)
    df = filter_by_seller_category_fixed(df, keys)
    print(f"📊 Rows after Seller Category filter: {len(df)}")
    if df.empty:
        print("❌ No data after applying Seller Category filter. Exiting.")
        return

    # separate seller data (on already category-filtered df)
    df_main, seller_dfs = separate_seller_data(df)

    # pick period
    start_date, end_date, period_name = get_date_range()

    # Filter main data
    df_cur = filter_by_date(df_main, start_date, end_date)
    print(f"📊 Main report: {len(df_cur)} rows in selected range.")
    if df_cur.empty:
        print("❌ No data in the chosen date range. Exiting.")
        return

    # Build previous period if needed for either executive deltas OR items comparison
    df_prev = None
    need_prev_for_exec = INCLUDE_PREVIOUS_PERIOD
    need_prev_for_items = INCLUDE_ITEMS_COMPARISON_SHEET
    if need_prev_for_exec or need_prev_for_items:
        days = (end_date - start_date).days + 1
        prev_start = start_date - timedelta(days=days)
        prev_end = start_date - timedelta(days=1)
        df_prev = filter_by_date(df_main, prev_start, prev_end)

    # Calculate main metrics
    calc = ReportMetrics(df_cur, df_prev if INCLUDE_PREVIOUS_PERIOD else None)
    metrics = calc.calculate_all()

    # Seller-specific metrics
    seller_analyses = {}
    if seller_dfs and INCLUDE_SELLER_ANALYSIS_SHEETS:
        print("\n" + "=" * 60)
        print("📊 ANALYZING SEPARATE SELLERS")
        print("=" * 60)
        for seller_name, seller_df in seller_dfs.items():
            seller_df_filtered = filter_by_date(seller_df, start_date, end_date)
            if not seller_df_filtered.empty:
                analyzer = SellerAnalysis(seller_name, seller_df_filtered)
                seller_analyses[seller_name] = analyzer.calculate_all()
            else:
                print(f"⚠️  No data for {seller_name} in selected date range")

    # output filename — include category label
    stamp = datetime.now().strftime("%Y%m%d") if INCLUDE_DATE_IN_FILENAME else ""
    safe_period = period_name.replace(" ", "_")
    safe_cats = (cats_label or "All").replace(" ", "_").replace(",", "+").replace("+", "_")
    out_name = f"{REPORT_NAME_PREFIX}_{safe_period}_{safe_cats}{('_' + stamp) if stamp else ''}.xlsx"

    # write
    generate_excel_report(
        metrics, seller_analyses,
        f"{period_name} — {cats_label}",
        out_name,
        df_cur,
        df_prev  # may be None; writer guards it
    )
    maybe_write_text_summary(metrics, f"{period_name} — {cats_label}", out_name[:-5])

    # summary to console
    ex = metrics["executive"]
    print("\n" + "=" * 60)
    print("✅ REPORT GENERATION COMPLETE")
    print("=" * 60)
    print(f"Categories : {cats_label}")
    print(f"Period     : {period_name}  ({start_date} → {end_date})")
    print(f"Revenue    : {ex['total_revenue']:,.0f} {CURRENCY_SYMBOL}")
    print(f"Orders     : {ex['total_orders']:,}")
    print(f"AOV        : {ex['avg_order_value']:,.0f} {CURRENCY_SYMBOL}")
    if INCLUDE_PREVIOUS_PERIOD:
        print(f"Rev Δ      : {ex.get('revenue_change', 0):.1f}%")
        print(f"Orders Δ   : {ex.get('orders_change', 0):.1f}%")
        if SHOW_ALERTS and ex.get("revenue_drop_alert"):
            print("⚠️  ALERT: Revenue down beyond threshold.")
        if SHOW_ALERTS and ex.get("orders_drop_alert"):
            print("⚠️  ALERT: Orders down beyond threshold.")

    if seller_analyses:
        print(f"\n📊 Separate Seller Analysis:")
        for seller_name, analysis in seller_analyses.items():
            if analysis and "summary" in analysis:
                s = analysis["summary"]
                print(f"   • {seller_name}: {s['total_orders']} orders, {s['total_revenue']:,.0f} {CURRENCY_SYMBOL}")

    print(f"\nSaved file : {out_name}\n")

if __name__ == "__main__":
    main()
