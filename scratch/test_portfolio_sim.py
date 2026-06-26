import sys, os
from datetime import datetime, date, timedelta
import pandas as pd

sys.path.insert(0, os.path.abspath("NAV_FETCHER"))
import nav_fetcher

# Test simulation with raw calculation using real AMFI data
def test_sim():
    print("--- Running test_portfolio_sim.py ---")
    bucket_df = pd.DataFrame([
        {"Scheme Name": "Quant Large Cap Fund-Reg(G)", "ISIN": "INF966L01AW4", "Weight (%)": 50.0},
        {"Scheme Name": "HDFC Flexi Cap Fund(G)", "ISIN": "INF179K01608", "Weight (%)": 50.0}
    ])
    
    start_date = date(2026, 6, 1)
    end_date = date(2026, 6, 10)
    initial_amount = 1000.0
    
    # Let's run raw calculations step-by-step
    bucket = bucket_df.copy()
    total_weight = bucket["Weight (%)"].sum()
    bucket["Weight_Normalized"] = bucket["Weight (%)"] / total_weight
    
    fetch_start = start_date - timedelta(days=10)
    isin_list = bucket["ISIN"].tolist()
    
    print("Fetching raw data...")
    df_raw = nav_fetcher.fetch_amfi_data_chunked(fetch_start, end_date, isin_list)
    print("df_raw count:", len(df_raw))
    
    # Robust date parsing
    parsed_dates = pd.to_datetime(df_raw["NAV Date"], format="%d-%b-%Y", errors="coerce")
    if parsed_dates.isna().all():
        parsed_dates = pd.to_datetime(df_raw["NAV Date"], errors="coerce")
    df_raw["NAV Date"] = parsed_dates
    df_raw = df_raw.dropna(subset=["NAV Date", "NAV"])
    
    df_raw["NAV_Date_Str"] = df_raw["NAV Date"].dt.strftime("%d-%m-%Y")
    
    df_pivot = df_raw.pivot_table(index="ISIN Div Payout / ISIN Growth", columns="NAV_Date_Str", values="NAV", aggfunc="first").reset_index()
    
    all_dates = nav_fetcher.build_date_cols(fetch_start, end_date, skip_sunday=False)
    for d in all_dates:
        if d not in df_pivot.columns:
            df_pivot[d] = None
            
    date_objs = sorted([datetime.strptime(d, "%d-%m-%Y") for d in all_dates])
    sorted_cols = [d.strftime("%d-%m-%Y") for d in date_objs]
    
    bucket["ISIN_upper"] = bucket["ISIN"].str.strip().str.upper()
    df_pivot["ISIN_upper"] = df_pivot["ISIN Div Payout / ISIN Growth"].str.strip().str.upper()
    
    df_sim = pd.merge(bucket, df_pivot, on="ISIN_upper", how="left")
    
    # Carry forward left-to-right
    for i in range(1, len(sorted_cols)):
        prev, curr = sorted_cols[i - 1], sorted_cols[i]
        df_sim[curr] = df_sim[curr].fillna(df_sim[prev])
        
    for i in range(len(sorted_cols) - 2, -1, -1):
        curr, next_col = sorted_cols[i], sorted_cols[i + 1]
        df_sim[curr] = df_sim[curr].fillna(df_sim[next_col])
        
    target_date_cols = nav_fetcher.build_date_cols(start_date, end_date, skip_sunday=True)
    t0 = target_date_cols[0]
    
    df_sim["Initial_Allocation"] = initial_amount * df_sim["Weight_Normalized"]
    df_sim["Units"] = df_sim["Initial_Allocation"] / df_sim[t0]
    
    daily_rows = []
    for d in target_date_cols:
        row_val = {"Date": d}
        total_val = 0.0
        for idx, row in df_sim.iterrows():
            name = row["Scheme Name"]
            nav_t = row[d]
            val_t = row["Units"] * nav_t
            row_val[f"{name} (₹)"] = round(val_t, 2)
            total_val += val_t
            
        row_val["Total Portfolio Value (₹)"] = round(total_val, 2)
        daily_rows.append(row_val)
        
    df_tracker = pd.DataFrame(daily_rows)
    print("df_tracker:")
    print(df_tracker.to_string())

if __name__ == "__main__":
    test_sim()
