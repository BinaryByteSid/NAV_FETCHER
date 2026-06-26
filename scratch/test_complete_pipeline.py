"""Test complete pipeline for app.py and nav_fetcher.py"""
import sys, os
from datetime import datetime, date
import pandas as pd

sys.path.insert(0, os.path.abspath("NAV_FETCHER"))

print("--- Testing NAV_FETCHER/app.py complete pipeline ---")
try:
    import app
    
    # Mock some data like download report would return
    df_port = app.load_portfolio_aum_data()
    
    # Let's mock a raw dataframe returned from AMFI download
    # 31-Mar-2026 (Tuesday) and 01-Apr-2026 (Wednesday)
    raw_rows = [
        {
            "Asset Class": "Open Ended Schemes(Equity Scheme - Flexi Cap Fund)",
            "Scheme Code": "103166",
            "ISIN Div Payout / ISIN Growth": "INF209K01AJ8",
            "ISIN Div Reinvestment": "-",
            "Scheme Name": "Aditya Birla Sun Life Flexi Cap Fund - Growth - Regular Plan",
            "NAV": 125.43,
            "Date": "31-Mar-2026"
        },
        {
            "Asset Class": "Open Ended Schemes(Equity Scheme - Flexi Cap Fund)",
            "Scheme Code": "103166",
            "ISIN Div Payout / ISIN Growth": "INF209K01AJ8",
            "ISIN Div Reinvestment": "-",
            "Scheme Name": "Aditya Birla Sun Life Flexi Cap Fund - Growth - Regular Plan",
            "NAV": 126.10,
            "Date": "01-Apr-2026"
        }
    ]
    df_raw = pd.DataFrame(raw_rows)
    df_raw["Plan Type"] = df_raw["Scheme Name"].apply(app.classify_plan_type)
    df_raw["Option Type"] = df_raw["Scheme Name"].apply(app.classify_option_type)
    
    # Run the exact sequence main() does:
    df_raw = app.populate_actual_aum(df_raw, df_port)
    parsed_dates = pd.to_datetime(df_raw["Date"], format="%d-%b-%Y", errors="coerce")
    if parsed_dates.isna().all():
        parsed_dates = pd.to_datetime(df_raw["Date"], errors="coerce")
    df_raw["Date"] = parsed_dates.dt.strftime("%d-%m-%Y")
    
    start_date = date(2026, 3, 31)
    end_date = date(2026, 4, 1)
    fetch_start_date = start_date
    skip_sundays = True
    carry_forward = True
    want_nav = True
    want_aum = True
    
    all_dates = pd.date_range(start=fetch_start_date, end=end_date)
    target_dates = []
    for dt in all_dates:
        if skip_sundays and dt.weekday() == 6:
            continue
        target_dates.append(dt.strftime("%d-%m-%Y"))
        
    fund_metadata = df_raw[[
        "Asset Class", "Scheme Code", "ISIN Div Payout / ISIN Growth", 
        "ISIN Div Reinvestment", "Scheme Name", "Plan Type", "Option Type"
    ]].drop_duplicates(subset=["Scheme Code"])
    
    df_pivot_nav = df_raw.pivot(index="Scheme Code", columns="Date", values="NAV").reset_index()
    df_pivot_aum = df_raw.pivot(index="Scheme Code", columns="Date", values="AUM").reset_index()
    
    nav_cols_map = {d: f"{d} (NAV)" for d in target_dates if d in df_pivot_nav.columns}
    aum_cols_map = {d: f"{d} (AUM)" for d in target_dates if d in df_pivot_aum.columns}
    
    df_pivot_nav = df_pivot_nav.rename(columns=nav_cols_map)
    df_pivot_aum = df_pivot_aum.rename(columns=aum_cols_map)
    df_pivot = pd.merge(df_pivot_nav, df_pivot_aum, on="Scheme Code", how="left")
    
    interleaved_dates = []
    for d in target_dates:
        interleaved_dates.append(f"{d} (NAV)")
        interleaved_dates.append(f"{d} (AUM)")
    display_date_cols = interleaved_dates
    
    df_final = pd.merge(fund_metadata, df_pivot, on="Scheme Code", how="left")
    for date_col in display_date_cols:
        if date_col not in df_final.columns:
            df_final[date_col] = None
            
    if carry_forward and len(target_dates) > 1:
        date_objs = sorted([datetime.strptime(d, "%d-%m-%Y") for d in target_dates])
        sorted_date_cols = [d.strftime("%d-%m-%Y") for d in date_objs]
        for i in range(1, len(sorted_date_cols)):
            prev_col = sorted_date_cols[i-1]
            curr_col = sorted_date_cols[i]
            df_final[f"{curr_col} (NAV)"] = df_final[f"{curr_col} (NAV)"].fillna(df_final[f"{prev_col} (NAV)"])
            df_final[f"{curr_col} (AUM)"] = df_final[f"{curr_col} (AUM)"].fillna(df_final[f"{prev_col} (AUM)"])
            
    if want_aum:
        df_pivot_fallback = df_raw.pivot(index="Scheme Code", columns="Date", values="Fallback_AUM").reset_index()
        fallback_cols_map = {f"{d} (AUM)": f"{d}_fallback_temp" for d in target_dates}
        df_pivot_fallback_renamed = df_pivot_fallback.rename(columns={d: f"{d}_fallback_temp" for d in target_dates if d in df_pivot_fallback.columns})
        available_temp_cols = [col for col in fallback_cols_map.values() if col in df_pivot_fallback_renamed.columns]
        df_final_temp = pd.merge(df_final, df_pivot_fallback_renamed[["Scheme Code"] + available_temp_cols], on="Scheme Code", how="left")
        for main_col, temp_col in fallback_cols_map.items():
            if main_col in df_final.columns and temp_col in df_final_temp.columns:
                df_final[main_col] = df_final[main_col].fillna(df_final_temp[temp_col])
                
    # Convert to vertical
    vertical_rows = []
    for _, row in df_final.iterrows():
        meta = {
            "Asset Class": row["Asset Class"],
            "Scheme Code": row["Scheme Code"],
            "ISIN Div Payout / ISIN Growth": row["ISIN Div Payout / ISIN Growth"],
            "ISIN Div Reinvestment": row["ISIN Div Reinvestment"],
            "Scheme Name": row["Scheme Name"],
            "Plan Type": row["Plan Type"],
            "Option Type": row["Option Type"]
        }
        for d in target_dates:
            r_item = meta.copy()
            r_item["NAV Date"] = d
            r_item["NAV"] = row[f"{d} (NAV)"]
            r_item["AUM Date"] = d
            r_item["AUM"] = row[f"{d} (AUM)"]
            vertical_rows.append(r_item)
            
    df_final_v = pd.DataFrame(vertical_rows)
    df_final_v = app.calculate_flows_for_dataframe(df_final_v, start_date.strftime("%d-%m-%Y"), list(fund_metadata.columns))
    print("[SUCCESS] app.py pipeline completed successfully!")
    print(df_final_v.to_string())
except Exception as e:
    print("[CRASH] app.py crashed with error:", e)
    import traceback
    traceback.print_exc()
