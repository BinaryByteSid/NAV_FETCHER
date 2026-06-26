import sys, os
from datetime import datetime, date, timedelta
import pandas as pd

sys.path.insert(0, os.path.abspath("NAV_FETCHER"))
import nav_fetcher

start_date = date(2026, 6, 1)
end_date = date(2026, 6, 10)
isin_list = ["INF200K01560"]

fetch_start_date = start_date - timedelta(days=10) # want_flows = True
df_raw = nav_fetcher.fetch_amfi_data_chunked(fetch_start_date, end_date, isin_list)

df_filtered = df_raw.copy()
print("Series elements types before converting:")
print(df_filtered["NAV Date"].apply(type).value_counts())

print("Attempting converting with pd.to_datetime:")
res = pd.to_datetime(df_filtered["NAV Date"], errors="coerce")
print(res)

print("Attempting with format='%d-%b-%Y':")
res_fmt = pd.to_datetime(df_filtered["NAV Date"], format="%d-%b-%Y", errors="coerce")
print(res_fmt)
