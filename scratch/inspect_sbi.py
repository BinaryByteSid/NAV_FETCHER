import sys, os
from datetime import datetime, date, timedelta
import pandas as pd

sys.path.insert(0, os.path.abspath("NAV_FETCHER"))
import nav_fetcher

start_date = date(2026, 6, 1)
end_date = date(2026, 6, 10)
isin_list = ["INF200K01560"]

fetch_start_date = start_date - timedelta(days=10) # want_flows = True
print(f"Fetching from {fetch_start_date} to {end_date}")
df_raw = nav_fetcher.fetch_amfi_data_chunked(fetch_start_date, end_date, isin_list)
print("df_raw columns:", df_raw.columns)
print("df_raw rows:")
print(df_raw[["Scheme Name", "NAV", "NAV Date"]].to_string())
