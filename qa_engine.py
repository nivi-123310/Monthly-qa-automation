import pandas as pd
import os
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta
import warnings

warnings.filterwarnings("ignore", category=UserWarning)

# ================= CONFIG =================

BASE_PATH = "data/input"
OUTPUT_PATH = "data/output"

os.makedirs(OUTPUT_PATH, exist_ok=True)

today = datetime.today()
target_date = today - relativedelta(months=1)

TARGET_MONTH = target_date.month
TARGET_YEAR = target_date.year
MONTH_NAME = target_date.strftime("%B %Y")

output_file = os.path.join(
    OUTPUT_PATH,
    f"QA_Report_{target_date.strftime('%B')}_{TARGET_YEAR}.xlsx"
)

# Generic markets (safe for GitHub)
VALID_MARKETS = ["Market_1", "Market_2", "Market_3"]

MANDATORY_COLUMNS = [
"Market","Media Channel","Media Sub Channel","Date","Month","Year",
"Brand","Product","Campaign","date_start","date_end",
"Currency","Spend Type","net_spend (local)","gross_spend (local)"
]

TEXT_COLUMNS = [
"Market","Media Channel","Media Sub Channel","Month",
"Brand","Product","Campaign","Currency","Spend Type"
]

# Generic mapping
BRAND_PRODUCT_MAP = {
    "Brand_A": ["Product_1", "Product_2"],
    "Brand_B": ["Product_3", "Product_4"]
}

print(f"🚀 QA Process Started for {MONTH_NAME}")
print("="*50)

# ================= HELPERS =================

def is_empty(val):
    if pd.isna(val):
        return True
    val_str = str(val).strip().lower()
    return val_str in ["", "na", "n/a", "none", "null"]

def clean_channel(folder):
    return folder.split(".")[-1].strip().upper()

def is_european_format(val):
    if is_empty(val):
        return False
    val = str(val).strip()
    pattern = r"^\d{1,3}(\.\d{3})*,\d{2}$"
    return bool(re.match(pattern, val))

def generate_file_reference(file):
    return f"file://{file}"

def is_valid_target(val):
    if pd.isna(val):
        return True
    try:
        if pd.notna(pd.to_datetime(val, errors='coerce')):
            return False
    except:
        pass
    return True

# ================= CAMPAIGN SPLIT =================

def detect_split_type(group):

    dates = sorted(pd.to_datetime(group['Date'], errors='coerce').dropna().unique())
    start = pd.to_datetime(group['date_start'].iloc[0], errors='coerce')
    end = pd.to_datetime(group['date_end'].iloc[0], errors='coerce')

    if pd.isna(start) or pd.isna(end):
        return "Irregular"

    duration = (end - start).days + 1
    distinct = len(dates)

    if duration == 1:
        return "Daily"

    if distinct == 1:
        return "Not Split"

    coverage = distinct / duration

    if coverage >= 0.95:
        return "Daily"

    if coverage >= 0.75:
        return "Partial (Few Days Missing)"

    if len(dates) > 1:
        gaps = [(dates[i+1] - dates[i]).days for i in range(len(dates)-1)]
        avg_gap = sum(gaps) / len(gaps)

        if 5 <= avg_gap <= 9:
            return "Weekly"

        if 25 <= avg_gap <= 35:
            return "Monthly"

    return "Irregular"

# ================= MAIN =================

all_issues = []
overview = []

for market in os.listdir(BASE_PATH):
    print(f"\n📊 Processing Market: {market}")

    if market not in VALID_MARKETS:
        continue

    market_path = os.path.join(BASE_PATH, market)

    for folder in os.listdir(market_path):

        channel_path = os.path.join(market_path, folder)
        if not os.path.isdir(channel_path):
            continue

        channel = clean_channel(folder)

        total_rows = 0
        issues_before = len(all_issues)

        full_data = []
        qa_month_data = []

        for file in os.listdir(channel_path):

            if not file.endswith(".xlsx"):
                continue

            file_path = os.path.join(channel_path, file)
            file_link = generate_file_reference(file)

            try:
                df = pd.read_excel(file_path, header=1, dtype=str)
                df = df.iloc[2:]
                df.columns = df.columns.str.strip()
                df = df.dropna(how='all')

                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

                full_data.append((df, file_link))

                qa_rows_file = []

                for _, row in df.iterrows():

                    present_count = sum(
                        0 if is_empty(row.get(col)) else 1
                        for col in MANDATORY_COLUMNS
                    )

                    if present_count == 0:
                        continue

                    date = row.get("Date")

                    if pd.notna(date):

                        if date.month == TARGET_MONTH and date.year == TARGET_YEAR:
                            qa_rows_file.append((row, file_link))

                        else:
                            start = pd.to_datetime(row.get("date_start"), errors='coerce')
                            end = pd.to_datetime(row.get("date_end"), errors='coerce')

                            month_start = pd.Timestamp(TARGET_YEAR, TARGET_MONTH, 1)
                            month_end = month_start + pd.offsets.MonthEnd(0)

                            if pd.notna(start) and pd.notna(end):
                                if start <= month_end and end >= month_start:
                                    qa_rows_file.append((row, file_link))

                    else:
                        qa_rows_file.append((row, file_link))

                if qa_rows_file:
                    qa_month_data.extend(qa_rows_file)
                    total_rows += len(qa_rows_file)

            except Exception as e:
                print(f"❌ Error in file {file}: {e}")
                continue

        if not qa_month_data:
            overview.append([market, channel, MONTH_NAME, "No Data", 0])
            continue

        full_df = pd.concat([d[0] for d in full_data], ignore_index=True)
        qa_rows = qa_month_data

        for row, file_link in qa_rows:

            campaign = row.get("Campaign") if not is_empty(row.get("Campaign")) else "Missing"
            sub = row.get("Media Sub Channel") if not is_empty(row.get("Media Sub Channel")) else "Missing"

            missing = [col for col in MANDATORY_COLUMNS if is_empty(row.get(col))]
            if missing:
                all_issues.append([market, channel, sub, campaign, "Missing Data", ", ".join(missing), file_link])

        status = "Pass" if len(all_issues) == issues_before else "Fail"
        overview.append([market, channel, MONTH_NAME, status, total_rows])

# ================= OUTPUT =================

qa_df = pd.DataFrame(all_issues, columns=[
"Market","Media Channel","Media Sub Channel","Campaign","Issue Type","Issue Details","File Link"
])

overview_df = pd.DataFrame(overview, columns=[
"Market","Media Channel","Month","Status","Rows Scanned"
])

with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    overview_df.to_excel(writer, sheet_name="Overview", index=False)
    qa_df.to_excel(writer, sheet_name="Issues", index=False)

print("✅ FINAL QA COMPLETED")
