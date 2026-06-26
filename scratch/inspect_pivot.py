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
df_filtered["NAV Date"] = pd.to_datetime(df_filtered["NAV Date"], errors="coerce")
df_filtered["NAV_Date_Str"] = df_filtered["NAV Date"].dt.strftime("%d-%m-%Y")

print("df_filtered before pivot:")
print(df_filtered[["Scheme Name", "NAV Date", "NAV_Date_Str", "NAV"]])

df_pivot_nav = df_filtered.pivot_table(index="Scheme Code", columns="NAV_Date_Str", values="NAV", aggfunc="first").reset_index()
print("df_pivot_nav columns:", df_pivot_nav.columns.tolist())
print("df_pivot_nav values:")
print(df_pivot_nav.to_string())

date_cols = nav_fetcher.build_date_cols(fetch_start_date, end_date, True)
print("date_cols:", date_cols)

nav_cols_map = {d: f"{d} (NAV)" for d in date_cols if d in df_pivot_nav.columns}
print("nav_cols_map:", nav_cols_map)
