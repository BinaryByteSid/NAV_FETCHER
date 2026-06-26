"""Full end-to-end simulation of Historical ISIN Export flow including Flows calculation"""
import requests
import sys, os
import pandas as pd
sys.path.insert(0, os.path.abspath("NAV_FETCHER"))
import app

# 1. Download exactly like the app does
url = "https://portal.amfiindia.com/DownloadNAVHistoryReport_Po.aspx?frmdt=30-Mar-2026&todt=02-Apr-2026"
headers = {"User-Agent": "Mozilla/5.0"}
resp = requests.get(url, headers=headers, timeout=60)

target_isins = {"INF209K01AJ8", "INF846K01CH7", "INF846K016E3", "INF194K01524",
                "INF179K01608", "INF179KA1RT1", "INF179K01CR2", "INF740K01128",
                "INF760K01019", "INF760K01KR2"}

# 2. Parse exactly like the app does (lines 1077-1125)
rows = []
current_section = "Unknown"
for line_bytes in resp.iter_lines():
    if not line_bytes:
        continue
    line = line_bytes.decode('utf-8', errors='ignore')
    if ";" not in line:
        line_stripped = line.strip()
        if (line_stripped.startswith("Open Ended") or line_stripped.startswith("Closed Ended") or line_stripped.startswith("Interval Fund Schemes")):
            current_section = line_stripped
        continue
    parts = line.split(";")
    if len(parts) < 8:
        continue
    isin_growth = parts[2].strip()
    isin_reinvestment = parts[3].strip()
    isin_growth_upper = isin_growth.upper() if isin_growth != "-" else ""
    isin_reinvest_upper = isin_reinvestment.upper() if isin_reinvestment != "-" else ""
    g_match = isin_growth_upper and isin_growth_upper in target_isins
    r_match = isin_reinvest_upper and isin_reinvest_upper in target_isins
    if g_match or r_match:
        scheme_code = parts[0].strip()
        scheme_name = parts[1].strip()
        nav_value = parts[4].strip()
        nav_date = parts[7].strip()
        scheme_code = scheme_code if scheme_code != "-" else None
        isin_growth = isin_growth if isin_growth != "-" else None
        isin_reinvestment = isin_reinvestment if isin_reinvestment != "-" else None
        nav = pd.to_numeric(nav_value.replace(",", ""), errors="coerce")
        rows.append({
            "Asset Class": current_section,
            "Scheme Code": scheme_code,
            "ISIN Div Payout / ISIN Growth": isin_growth,
            "ISIN Div Reinvestment": isin_reinvestment,
            "Scheme Name": scheme_name,
            "NAV": nav,
            "Date": nav_date
        })

df_raw = pd.DataFrame(rows)
print(f"Parsed {len(df_raw)} rows from historical download\n")

# 3. Run populate_actual_aum  
df_port = app.load_portfolio_aum_data()
df_res = app.populate_actual_aum(df_raw, df_port)

# Now, we do pivot and fill NaNs like the app does in app.py main()
# (simulating app.py lines 1142-1282)
target_dates = sorted(list(df_res["Date"].unique()))
print(f"Target dates: {target_dates}\n")

# Pivot df_res to get the wide format for final output
# We pivot NAVs and AUMs
df_pivot_nav = df_res.pivot(index="Scheme Code", columns="Date", values="NAV").reset_index()
df_pivot_aum = df_res.pivot(index="Scheme Code", columns="Date", values="AUM").reset_index()

# Rename columns
df_pivot_nav = df_pivot_nav.rename(columns={d: f"{d} (NAV)" for d in target_dates})
df_pivot_aum = df_pivot_aum.rename(columns={d: f"{d} (AUM)" for d in target_dates})

# Merge metadata
df_meta = df_res[["Asset Class", "Scheme Code", "ISIN Div Payout / ISIN Growth", "ISIN Div Reinvestment", "Scheme Name"]].drop_duplicates(subset=["Scheme Code"])
df_final = pd.merge(df_meta, df_pivot_nav, on="Scheme Code", how="left")
df_final = pd.merge(df_final, df_pivot_aum, on="Scheme Code", how="left")

# Add missing columns
for d in target_dates:
    if f"{d} (NAV)" not in df_final.columns:
        df_final[f"{d} (NAV)"] = None
    if f"{d} (AUM)" not in df_final.columns:
        df_final[f"{d} (AUM)"] = None

# Fill NaNs using carry-forward like app.py does:
# (app.py lines 1197-1208)
for i in range(1, len(target_dates)):
    prev_d = target_dates[i-1]
    curr_d = target_dates[i]
    df_final[f"{curr_d} (NAV)"] = df_final[f"{curr_d} (NAV)"].fillna(df_final[f"{prev_d} (NAV)"])
    df_final[f"{curr_d} (AUM)"] = df_final[f"{curr_d} (AUM)"].fillna(df_final[f"{prev_d} (AUM)"])

# Fill any remaining NaNs in AUM columns with fallback values:
# (app.py lines 1210-1226)
df_pivot_fallback = df_res.pivot(index="Scheme Code", columns="Date", values="Fallback_AUM").reset_index()
fallback_cols_map = {f"{d} (AUM)": f"{d}_fallback_temp" for d in target_dates}
df_pivot_fallback_renamed = df_pivot_fallback.rename(columns={d: f"{d}_fallback_temp" for d in target_dates if d in df_pivot_fallback.columns})
available_temp_cols = [col for col in fallback_cols_map.values() if col in df_pivot_fallback_renamed.columns]
df_final_temp = pd.merge(df_final, df_pivot_fallback_renamed[["Scheme Code"] + available_temp_cols], on="Scheme Code", how="left")
for main_col, temp_col in fallback_cols_map.items():
    if main_col in df_final.columns and temp_col in df_final_temp.columns:
        df_final[main_col] = df_final[main_col].fillna(df_final_temp[temp_col])

# Convert df_final to vertical layout like app.py does:
# (app.py lines 1228-1253)
vertical_rows = []
for _, row in df_final.iterrows():
    meta = {
        "Asset Class": row["Asset Class"],
        "Scheme Code": row["Scheme Code"],
        "ISIN Div Payout / ISIN Growth": row["ISIN Div Payout / ISIN Growth"],
        "ISIN Div Reinvestment": row["ISIN Div Reinvestment"],
        "Scheme Name": row["Scheme Name"]
    }
    for d in target_dates:
        r_item = meta.copy()
        r_item["NAV Date"] = d
        r_item["NAV"] = row[f"{d} (NAV)"]
        r_item["AUM Date"] = d
        r_item["AUM"] = row[f"{d} (AUM)"]
        vertical_rows.append(r_item)

df_vertical = pd.DataFrame(vertical_rows)

# Calculate flows!
# (app.py line 1283)
start_date = "31-Mar-2026"
df_flows = app.calculate_flows_for_dataframe(df_vertical, start_date, ["Asset Class", "Scheme Code", "ISIN Div Payout / ISIN Growth", "ISIN Div Reinvestment", "Scheme Name"])

print("Flows calculation results:")
print(df_flows[["Scheme Name", "NAV Date", "NAVs", "Closing AUM as on previous day", "Actual AUM as on current date", "Daily return", "Derived AUM as on curent day", "Net flows on current day"]].to_string())
