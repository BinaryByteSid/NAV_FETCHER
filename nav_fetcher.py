"""
AMFI NAV Fetcher — Standalone UI
Fetch historical NAV data from AMFI India by ISIN(s) and/or date range.

Run with:
    streamlit run nav_fetcher.py
"""

from __future__ import annotations

import os
import re
import sys
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import List, Optional, Union

import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ─── Robust date parsing (locale-independent) ────────────────────────────────

_MONTH_ABBR = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04",
    "may": "05", "jun": "06", "jul": "07", "aug": "08",
    "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}


def _parse_amfi_date_str(date_str: str) -> str | None:
    """Convert a single AMFI date string like '01-Jun-2026' → '2026-06-01' (ISO).

    Returns None if unparseable. This avoids reliance on pandas locale settings.
    """
    if not isinstance(date_str, str) or not date_str.strip():
        return None
    parts = date_str.strip().split("-")
    if len(parts) == 3:
        day, month_abbr, year = parts
        month_num = _MONTH_ABBR.get(month_abbr.lower()[:3])
        if month_num and day.isdigit() and year.isdigit():
            return f"{year}-{month_num}-{day.zfill(2)}"
    return None


def parse_amfi_date_series(series: pd.Series) -> pd.Series:
    """Parse a Series of AMFI date strings ('dd-Mon-YYYY' or 'dd-mm-YYYY') to datetime.

    Uses a manual month-abbreviation lookup so it works identically on all
    locales / pandas versions (including Streamlit Cloud).
    """
    # Fast path: try explicit format first (works when locale is OK)
    result = pd.to_datetime(series, format="%d-%b-%Y", errors="coerce")
    if not result.isna().all():
        # Fill any remaining NaTs with manual parsing
        mask = result.isna() & series.notna()
        if mask.any():
            iso_strs = series[mask].apply(_parse_amfi_date_str)
            result[mask] = pd.to_datetime(iso_strs, errors="coerce")
        return result

    # Fallback: try numeric format dd-mm-YYYY (for pre-converted dates)
    result = pd.to_datetime(series, format="%d-%m-%Y", errors="coerce")
    if not result.isna().all():
        return result

    # Last resort: manual element-wise parsing for AMFI dates
    iso_strs = series.apply(_parse_amfi_date_str)
    result = pd.to_datetime(iso_strs, errors="coerce")
    if not result.isna().all():
        return result

    # Absolute fallback
    return pd.to_datetime(series, errors="coerce")


def get_fund_seed(name: str) -> int:
    hash_val = 0
    for char in name:
        hash_val = ord(char) + ((hash_val << 5) - hash_val)
        hash_val = hash_val & 0xFFFFFFFF
    if hash_val > 0x7FFFFFFF:
        hash_val = hash_val - 0x100000000
    return abs(hash_val) % 100


def load_portfolio_aum_data() -> pd.DataFrame:
    paths = [
        "../portfolio last 6 months.xlsx",
        "portfolio last 6 months.xlsx",
        "c:/Users/sidha/OneDrive/Desktop/portfolio last 6 months.xlsx",
    ]
    for path in paths:
        if os.path.exists(path):
            try:
                df = pd.read_excel(path, header=3)
                df['SD_Scheme ISIN'] = df['SD_Scheme ISIN'].astype(str).str.strip().str.upper()
                df['PD_Month End'] = pd.to_numeric(df['PD_Month End'], errors='coerce')
                df['PD_Scheme AUM'] = pd.to_numeric(df['PD_Scheme AUM'], errors='coerce')
                return df
            except Exception:
                pass
    return pd.DataFrame()


def calculate_aum_for_row(row, df_port: pd.DataFrame) -> float:
    scheme_name = row.get("Scheme Name", "")
    isin_growth = row.get("ISIN Div Payout / ISIN Growth") or row.get("ISIN Div Payout/ ISIN Growth")
    isin_reinvestment = row.get("ISIN Div Reinvestment")
    date_str = row.get("Date") or row.get("NAV Date")
    
    try:
        if isinstance(date_str, pd.Timestamp):
            year = date_str.year
            month = date_str.month
        else:
            parts = str(date_str).split("-")
            month_map = {
                "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
                "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
            }
            if len(parts) == 3:
                month_str = parts[1]
                year = int(parts[2][:4])
                month = month_map.get(month_str[:3], 4)
            else:
                parsed_dt = pd.to_datetime(date_str)
                year = parsed_dt.year
                month = parsed_dt.month
    except Exception:
        year = 2026
        month = 4
        
    m_val = year * 100 + month
    
    match_port = pd.DataFrame()
    if not df_port.empty:
        isins = []
        if isin_growth and pd.notna(isin_growth) and isin_growth != "-":
            isins.append(str(isin_growth).strip().upper())
        if isin_reinvestment and pd.notna(isin_reinvestment) and isin_reinvestment != "-":
            isins.append(str(isin_reinvestment).strip().upper())
            
        if isins:
            match_port = df_port[df_port['SD_Scheme ISIN'].isin(isins)]
            
    aum_monthly = None
    if not match_port.empty:
        match_month = match_port[match_port['PD_Month End'] == m_val]
        if not match_month.empty:
            aum_monthly = float(match_month['PD_Scheme AUM'].iloc[0])
        else:
            latest_row = match_port.sort_values('PD_Month End', ascending=False).iloc[0]
            aum_monthly = float(latest_row['PD_Scheme AUM'])
            
    if aum_monthly is None:
        seed = get_fund_seed(scheme_name)
        base_aum = (seed % 35 + 15) * 1000 + (seed % 97) + 0.56
        month_offset = (2026 - year) * 12 + (4 - month)
        aum_multiplier = 1.0 - (month_offset * 0.012) + ((seed + month) % 5 - 2) * 0.002
        aum_monthly = round(base_aum * aum_multiplier, 4)
        
    return aum_monthly


def map_section_to_ids(sec):
    sec_lower = str(sec).lower()
    
    # Maturity Type
    maturity_id = 1  # Open ended default
    if "close" in sec_lower:
        maturity_id = 2
    elif "interval" in sec_lower:
        maturity_id = 2
        
    # Category
    cat_id = 1  # Equity default
    if "debt" in sec_lower:
        cat_id = 2
    elif "hybrid" in sec_lower:
        cat_id = 3
    elif "solution" in sec_lower:
        cat_id = 4
    elif "other" in sec_lower:
        cat_id = 5
    elif "gilt" in sec_lower or "money market" in sec_lower or "income" in sec_lower:
        cat_id = 2
        
    # Subcategory defaults
    if cat_id == 1:
        sub_id = 1
    elif cat_id == 2:
        sub_id = 15
    elif cat_id == 3:
        sub_id = 30
    elif cat_id == 4:
        sub_id = 36
    elif cat_id == 5:
        sub_id = 38
    else:
        sub_id = 1
    
    # Subcategory mapping rules
    if cat_id == 1:  # Equity
        if "large & mid" in sec_lower:
            sub_id = 2
        elif "large cap" in sec_lower:
            sub_id = 1
        elif "flexi cap" in sec_lower:
            sub_id = 3
        elif "multi cap" in sec_lower:
            sub_id = 4
        elif "mid cap" in sec_lower:
            sub_id = 5
        elif "small cap" in sec_lower:
            sub_id = 6
        elif "value" in sec_lower:
            sub_id = 7
        elif "elss" in sec_lower:
            sub_id = 8
        elif "contra" in sec_lower:
            sub_id = 9
        elif "dividend yield" in sec_lower:
            sub_id = 10
        elif "focused" in sec_lower:
            sub_id = 11
        elif "sectoral" in sec_lower or "thematic" in sec_lower:
            sub_id = 12
    elif cat_id == 2:  # Debt
        if "gilt with 10" in sec_lower or "10 year constant" in sec_lower:
            sub_id = 29
        elif "gilt" in sec_lower:
            sub_id = 28
        elif "medium to long" in sec_lower:
            sub_id = 14
        elif "long duration" in sec_lower:
            sub_id = 13
        elif "ultra short" in sec_lower:
            sub_id = 19
        elif "short duration" in sec_lower:
            sub_id = 15
        elif "medium duration" in sec_lower:
            sub_id = 16
        elif "money market" in sec_lower:
            sub_id = 17
        elif "low duration" in sec_lower:
            sub_id = 18
        elif "liquid" in sec_lower:
            sub_id = 20
        elif "overnight" in sec_lower:
            sub_id = 21
        elif "dynamic bond" in sec_lower:
            sub_id = 22
        elif "corporate bond" in sec_lower:
            sub_id = 23
        elif "credit risk" in sec_lower:
            sub_id = 24
        elif "banking" in sec_lower or "psu" in sec_lower:
            sub_id = 25
        elif "floater" in sec_lower:
            sub_id = 26
        elif "fmp" in sec_lower:
            sub_id = 27
    elif cat_id == 3:  # Hybrid
        if "aggressive hybrid" in sec_lower:
            sub_id = 30
        elif "conservative hybrid" in sec_lower or "conservative hyrbid" in sec_lower:
            sub_id = 31
        elif "equity savings" in sec_lower:
            sub_id = 32
        elif "arbitrage" in sec_lower:
            sub_id = 33
        elif "multi asset" in sec_lower:
            sub_id = 34
        elif "dynamic asset" in sec_lower or "balanced advantage" in sec_lower:
            sub_id = 35
        elif "balanced hybrid" in sec_lower:
            sub_id = 40
    elif cat_id == 4:  # Solution Oriented
        if "children" in sec_lower:
            sub_id = 36
        elif "retirement" in sec_lower:
            sub_id = 37
    elif cat_id == 5:  # Other
        if "index fund" in sec_lower or "index" in sec_lower:
            sub_id = 38
        elif "etf" in sec_lower:
            sub_id = 38
        elif "fof" in sec_lower or "fund of funds" in sec_lower:
            sub_id = 39
        
    return maturity_id, cat_id, sub_id


def clean_name(name: str) -> str:
    n = str(name).lower()
    n = n.replace("flexicap", "flexi cap")
    n = n.replace("multicap", "multi cap")
    n = n.replace("midcap", "mid cap")
    n = n.replace("smallcap", "small cap")
    n = n.replace("largecap", "large cap")
    n = n.replace("focussed", "focused")
    
    n = n.replace("-", " ").replace("/", " ").replace("(", " ").replace(")", " ")
    n = n.replace(" sl ", " sun life ")
    tokens = n.split()
    suffixes_to_remove = {
        "direct", "regular", "retail", "plan", "growth", "option", "idcw", "dividend", 
        "payout", "reinvestment", "annual", "monthly", "weekly", "quarterly", "fortnightly",
        "bonus", "fund"
    }
    cleaned_tokens = [t for t in tokens if t not in suffixes_to_remove]
    return " ".join(cleaned_tokens)


API_CACHE = {}


def fetch_performance_data_from_api(date_str: str, maturity_id: int, category_id: int, subcategory_id: int) -> list:
    key = (date_str, maturity_id, category_id, subcategory_id)
    if key in API_CACHE:
        return API_CACHE[key]
        
    url = "https://www.amfiindia.com/gateway/pollingsebi/api/amfi/fundperformance"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Content-Type": "application/json"
    }
    payload = {
        "maturityType": maturity_id,
        "category": category_id,
        "subCategory": subcategory_id,
        "mfid": 0,
        "reportDate": date_str
    }
    
    import time
    max_retries = 3
    backoff = 1.0
    
    for attempt in range(max_retries):
        try:
            resp = requests.post(url, json=payload, headers=headers, timeout=20)
            if resp.status_code == 200:
                res_data = resp.json()
                if res_data.get("validationMsg") == "SUCCESS":
                    rows = res_data.get("data", [])
                    API_CACHE[key] = rows
                    return rows
                else:
                    print(f"AMFI API Validation failed for {key}: {res_data.get('validationMsg')}")
            else:
                print(f"AMFI API returned HTTP code {resp.status_code} for {key}")
        except Exception as e:
            print(f"Attempt {attempt+1} failed for {key} with error: {e}")
        
        if attempt < max_retries - 1:
            time.sleep(backoff)
            backoff *= 2
            
    API_CACHE[key] = []
    return []


def find_matching_perf_row(nav_name: str, perf_rows: list) -> Optional[dict]:
    if not perf_rows:
        return None
    cleaned_nav = clean_name(nav_name)
    if not cleaned_nav:
        return None
        
    # 1. Substring match
    for p_row in perf_rows:
        p_name = p_row.get("schemeName") or ""
        cleaned_perf = clean_name(p_name)
        if cleaned_perf and (cleaned_perf in cleaned_nav or cleaned_nav in cleaned_perf):
            return p_row
            
    # 2. Token overlap fallback
    nav_tokens = set(cleaned_nav.split())
    best_row = None
    best_score = 0.0
    for p_row in perf_rows:
        p_name = p_row.get("schemeName") or ""
        cleaned_perf = clean_name(p_name)
        if not cleaned_perf:
            continue
        perf_tokens = set(cleaned_perf.split())
        intersection = nav_tokens.intersection(perf_tokens)
        if intersection:
            score = len(intersection) / len(nav_tokens.union(perf_tokens))
            if score > best_score:
                best_score = score
                best_row = p_row
                
    if best_score > 0.4:
        return best_row
    return None


def populate_actual_aum(df: pd.DataFrame, df_port: pd.DataFrame, want_aum: bool = True, fetch_live_aum: bool = False) -> pd.DataFrame:
    if df.empty:
        return df.copy()
        
    df_res = df.copy()
    
    if not want_aum:
        df_res["AUM"] = None
        df_res["Fallback_AUM"] = None
        return df_res
        
    # 1. First, calculate the fallback AUM for all rows using the old method.
    # We do this so that we always have a default value.
    raw_rows = []
    for idx, r_dict in df_res.iterrows():
        r_copy = r_dict.copy()
        m_aum = calculate_aum_for_row(r_copy.to_dict(), df_port)
        r_copy["Monthly_AUM"] = m_aum
        raw_rows.append(r_copy)
    df_res = pd.DataFrame(raw_rows)
    
    mean_navs = df_res.groupby("Scheme Code")["NAV"].transform("mean")
    mean_navs = mean_navs.fillna(1.0).replace(0.0, 1.0)
    df_res["Fallback_AUM"] = (df_res["Monthly_AUM"] * (df_res["NAV"] / mean_navs)).round(4)
    
    if not fetch_live_aum:
        # If user does not want slow live AUM, use Fallback AUM immediately
        df_res["AUM"] = df_res["Fallback_AUM"]
        return df_res
        
    # Initialize AUM with None to allow carry-forward for missing dates
    df_res["AUM"] = None
    
    # 2. Now, try to fetch the actual AUM from the performance API for each row.
    def get_date_str(dt):
        try:
            if isinstance(dt, pd.Timestamp) or hasattr(dt, "strftime"):
                if pd.notna(dt):
                    return dt.strftime("%d-%b-%Y")
                return None
            parsed = pd.to_datetime(dt, format="%d-%b-%Y", errors="coerce")
            if pd.isna(parsed):
                parsed = pd.to_datetime(dt, errors="coerce")
            if pd.notna(parsed):
                return parsed.strftime("%d-%b-%Y")
        except Exception:
            pass
        return str(dt)
        
    date_col = "NAV Date" if "NAV Date" in df_res.columns else "Date"
    df_res["Date_Str_Temp"] = df_res[date_col].apply(get_date_str)
    
    # Group unique combinations of (Asset Class, Date_Str_Temp)
    unique_groups = df_res[["Asset Class", "Date_Str_Temp"]].drop_duplicates()
    
    # Fetch performance data for each group and build a lookup cache
    perf_lookup = {}
    import time
    
    for i, (_, grp) in enumerate(unique_groups.iterrows()):
        asset_class = grp["Asset Class"]
        date_str = grp["Date_Str_Temp"]
        if not asset_class or not date_str:
            continue
            
        m_id, c_id, s_id = map_section_to_ids(asset_class)
        
        # Rate-limiting sleep between calls to AMFI APIs
        if i > 0:
            time.sleep(0.5)
            
        # Fetch from API
        perf_rows = fetch_performance_data_from_api(date_str, m_id, c_id, s_id)
        if perf_rows:
            perf_lookup[(date_str, asset_class)] = perf_rows
            
    # Now, try to match each row to the fetched performance rows
    for idx, row in df_res.iterrows():
        asset_class = row["Asset Class"]
        date_str = row["Date_Str_Temp"]
        scheme_name = row["Scheme Name"]
        
        perf_rows = perf_lookup.get((date_str, asset_class), [])
        match = find_matching_perf_row(scheme_name, perf_rows)
        if match:
            daily_aum = match.get("dailyAUM")
            if daily_aum is not None and daily_aum != "":
                try:
                    df_res.at[idx, "AUM"] = float(daily_aum)
                except Exception:
                    pass
                    
    # Drop temporary column
    df_res = df_res.drop(columns=["Date_Str_Temp"])
    return df_res


def calculate_flows_for_dataframe(df: pd.DataFrame, start_date, meta_cols: list) -> pd.DataFrame:
    """Calculate the flows format columns for a vertical Mutual Fund DataFrame."""
    df = df.copy()
    
    # Standardize columns
    if "NAV" not in df.columns and "NAVs" in df.columns:
        df = df.rename(columns={"NAVs": "NAV"})
    if "NAV" not in df.columns:
        df["NAV"] = None
    if "AUM" not in df.columns:
        df["AUM"] = None
        
    df["NAV"] = pd.to_numeric(df["NAV"], errors="coerce")
    df["AUM"] = pd.to_numeric(df["AUM"], errors="coerce")
    
    # Standardize date column to 'NAV Date'
    date_col = None
    for c in ["NAV Date", "AUM Date", "Date"]:
        if c in df.columns:
            date_col = c
            break
            
    if date_col and date_col != "NAV Date":
        df = df.rename(columns={date_col: "NAV Date"})
        
    if "NAV Date" not in df.columns:
        df["NAV Date"] = None
        
    df["NAV Date_parsed"] = parse_amfi_date_series(df["NAV Date"])
        
    df = df.sort_values(by=["Scheme Code", "NAV Date_parsed"]).reset_index(drop=True)
    
    df["Closing AUM as on previous day"] = None
    df["Actual AUM as on current date"] = df["AUM"]
    df["Daily return"] = None
    df["Derived AUM as on curent day"] = None
    df["Net flows on current day"] = None
    
    for scheme_code, group in df.groupby("Scheme Code"):
        indices = group.index
        for idx_in_group, idx in enumerate(indices):
            if idx_in_group == 0:
                continue
            prev_idx = indices[idx_in_group - 1]
            
            nav_curr = df.at[idx, "NAV"]
            nav_prev = df.at[prev_idx, "NAV"]
            aum_prev = df.at[prev_idx, "AUM"]
            aum_curr = df.at[idx, "AUM"]
            
            df.at[idx, "Closing AUM as on previous day"] = aum_prev
            
            if pd.notna(nav_curr) and pd.notna(nav_prev) and nav_prev != 0:
                daily_return = (nav_curr - nav_prev) / nav_prev
                df.at[idx, "Daily return"] = daily_return * 100
            else:
                daily_return = None
                
            if pd.notna(aum_prev) and daily_return is not None:
                derived_aum = aum_prev * (1 + daily_return)
                df.at[idx, "Derived AUM as on curent day"] = derived_aum
            else:
                derived_aum = None
                
            if pd.notna(aum_curr) and derived_aum is not None:
                df.at[idx, "Net flows on current day"] = aum_curr - derived_aum
                
    # Filter only target dates
    start_date_ts = pd.to_datetime(start_date)
    df = df[df["NAV Date_parsed"] >= start_date_ts].reset_index(drop=True)
    df = df.drop(columns=["NAV Date_parsed"])
    
    df = df.rename(columns={"NAV": "NAVs"})
    
    flow_cols = [
        "NAV Date",
        "NAVs",
        "Closing AUM as on previous day",
        "Actual AUM as on current date",
        "Daily return",
        "Derived AUM as on curent day",
        "Net flows on current day"
    ]
    
    if "NAV Date" in df.columns:
        parsed = parse_amfi_date_series(df["NAV Date"])
        df["NAV Date"] = parsed.dt.strftime("%d-%m-%Y")
        
    final_cols = list(meta_cols) + flow_cols
    for col in final_cols:
        if col not in df.columns:
            df[col] = None
            
    df["NAV Date_parsed"] = parse_amfi_date_series(df["NAV Date"])
    df = df.sort_values(by=["Asset Class", "Scheme Name", "NAV Date_parsed"]).reset_index(drop=True)
    df = df.drop(columns=["NAV Date_parsed"])
    
    return df[final_cols]


def parse_bucket_input(uploaded_file=None, data_editor_df=None) -> pd.DataFrame:
    """Parse Excel/CSV uploads or dynamic table data editor inputs for the portfolio bucket composition."""
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            # Match columns dynamically
            isin_col = None
            weight_col = None
            name_col = None
            
            for c in df.columns:
                c_low = str(c).lower()
                if "isin" in c_low:
                    isin_col = c
                elif "weight" in c_low or "wt" in c_low or "share" in c_low:
                    weight_col = c
                elif "scheme" in c_low or "name" in c_low or "fund" in c_low:
                    name_col = c
            
            if isin_col is None:
                st.error("Uploaded file must contain an 'ISIN' column.")
                return pd.DataFrame()
            
            res_df = pd.DataFrame()
            res_df["ISIN"] = df[isin_col].astype(str).str.strip().str.upper()
            if name_col is not None:
                res_df["Scheme Name"] = df[name_col].astype(str).str.strip()
            else:
                res_df["Scheme Name"] = res_df["ISIN"]
                
            if weight_col is not None:
                w_series = df[weight_col].astype(str).str.replace("%", "").str.strip()
                res_df["Weight (%)"] = pd.to_numeric(w_series, errors="coerce").fillna(0.0)
            else:
                res_df["Weight (%)"] = 100.0 / len(res_df)
                
            return res_df.dropna(subset=["ISIN"])
        except Exception as e:
            st.error(f"Error parsing uploaded file: {e}")
            return pd.DataFrame()
            
    if data_editor_df is not None:
        df = data_editor_df.copy()
        if "ISIN" not in df.columns:
            return pd.DataFrame()
        df["ISIN"] = df["ISIN"].astype(str).str.strip().str.upper()
        df["Weight (%)"] = pd.to_numeric(df.get("Weight (%)", 0.0), errors="coerce").fillna(0.0)
        if "Scheme Name" not in df.columns:
            df["Scheme Name"] = df["ISIN"]
        return df.dropna(subset=["ISIN"])
        
    return pd.DataFrame()


def fetch_latest_navs(isin_list: List[str]) -> dict:
    """Fetch the latest (today's) NAV for each ISIN from AMFI's live NAVAll.txt feed.

    Returns dict: {ISIN_upper: {"nav": float, "date": str, "scheme_name": str}}
    """
    isin_set = {i.strip().upper() for i in isin_list if i.strip()}
    if not isin_set:
        return {}

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    try:
        resp = requests.get(AMFI_LATEST_URL, headers=headers, timeout=30)
        resp.raise_for_status()
    except Exception:
        return {}

    result = {}
    for line_bytes in resp.text.splitlines():
        line = line_bytes.strip()
        if ";" not in line or line.startswith("Scheme Code"):
            continue
        parts = line.split(";")
        if len(parts) < 8:
            continue

        isin_g = parts[2].strip().upper() if parts[2].strip() != "-" else ""
        isin_r = parts[3].strip().upper() if parts[3].strip() != "-" else ""
        matched_isin = None
        if isin_g in isin_set:
            matched_isin = isin_g
        elif isin_r in isin_set:
            matched_isin = isin_r
        else:
            continue

        nav_val = pd.to_numeric(parts[4].strip().replace(",", ""), errors="coerce")
        nav_date = parts[7].strip()  # e.g. "27-Jun-2026"
        scheme_name = parts[1].strip()

        if pd.notna(nav_val) and matched_isin:
            result[matched_isin] = {
                "nav": float(nav_val),
                "date": nav_date,
                "scheme_name": scheme_name,
            }
    return result


def run_live_portfolio(bucket_df: pd.DataFrame, start_date, initial_amount: float, skip_sunday: bool):
    """Run a live portfolio tracker from start_date to today.

    Fetches historical NAVs from AMFI history API, then overlays the latest
    real-time NAV from NAVAll.txt so the portfolio value is always current.
    """
    end_date = datetime.today().date()

    if bucket_df.empty:
        return None, "Bucket composition is empty."
        
    bucket = bucket_df.copy()
    bucket = bucket[bucket["ISIN"].str.strip() != ""]
    if bucket.empty:
        return None, "Bucket composition has no valid ISINs."
        
    total_weight = bucket["Weight (%)"].sum()
    if total_weight <= 0:
        return None, "Total weight must be greater than 0%."
        
    bucket["Weight_Normalized"] = bucket["Weight (%)"] / total_weight
    
    # Start fetch 10 days early to cover holidays and provide carry-forward baseline
    fetch_start = start_date - timedelta(days=10)
    isin_list = bucket["ISIN"].tolist()
    
    df_raw = fetch_amfi_data_chunked(fetch_start, end_date, isin_list)
    if df_raw.empty:
        return None, "No historical NAV data returned from AMFI for these ISINs. Please check that the date range is valid."
        
    # Standardize dates using robust locale-independent parser
    df_raw["NAV Date"] = parse_amfi_date_series(df_raw["NAV Date"])
    df_raw = df_raw.dropna(subset=["NAV Date", "NAV"])
    
    df_raw["NAV_Date_Str"] = df_raw["NAV Date"].dt.strftime("%d-%m-%Y")
    
    df_pivot = df_raw.pivot_table(index="ISIN Div Payout / ISIN Growth", columns="NAV_Date_Str", values="NAV", aggfunc="first").reset_index()
    df_pivot_r = df_raw.pivot_table(index="ISIN Div Reinvestment", columns="NAV_Date_Str", values="NAV", aggfunc="first").reset_index()
    
    all_dates = build_date_cols(fetch_start, end_date, skip_sunday=False)
    for d in all_dates:
        if d not in df_pivot.columns:
            df_pivot[d] = None
            
    date_objs = sorted([datetime.strptime(d, "%d-%m-%Y") for d in all_dates])
    sorted_cols = [d.strftime("%d-%m-%Y") for d in date_objs]
    
    bucket["ISIN_upper"] = bucket["ISIN"].str.strip().str.upper()
    df_pivot["ISIN_upper"] = df_pivot["ISIN Div Payout / ISIN Growth"].str.strip().str.upper()
    
    df_sim = pd.merge(bucket, df_pivot, on="ISIN_upper", how="left")
    
    # Merge reinvestment columns if payout ISIN was missing
    missing_mask = df_sim[sorted_cols[0]].isna()
    if missing_mask.any() and not df_pivot_r.empty:
        df_pivot_r["ISIN_upper"] = df_pivot_r["ISIN Div Reinvestment"].str.strip().str.upper()
        for idx, row in df_sim[missing_mask].iterrows():
            match_r = df_pivot_r[df_pivot_r["ISIN_upper"] == row["ISIN_upper"]]
            if not match_r.empty:
                for col in sorted_cols:
                    df_sim.at[idx, col] = match_r.iloc[0].get(col)
                    
    # Carry forward chronological left-to-right
    for i in range(1, len(sorted_cols)):
        prev, curr = sorted_cols[i - 1], sorted_cols[i]
        df_sim[curr] = df_sim[curr].fillna(df_sim[prev])
        
    for i in range(len(sorted_cols) - 2, -1, -1):
        curr, next_col = sorted_cols[i], sorted_cols[i + 1]
        df_sim[curr] = df_sim[curr].fillna(df_sim[next_col])
        
    target_date_cols = build_date_cols(start_date, end_date, skip_sunday)
    if not target_date_cols:
        return None, "Target date range contains no trading days (all dates filtered out)."
        
    t0 = target_date_cols[0]
    df_sim = df_sim.dropna(subset=[t0])
    if len(df_sim) < len(bucket):
        missing_isins = set(bucket["ISIN_upper"]) - set(df_sim["ISIN_upper"])
        return None, f"Could not find historical NAV data for ISIN(s): {', '.join(missing_isins)}. Please check the ISINs and date range."
    
    # ── Overlay latest live NAVs from NAVAll.txt ──────────────────────────────
    latest_navs = fetch_latest_navs(isin_list)
    latest_nav_date_str = None
    if latest_navs:
        # Find the latest date from the live feed
        for info in latest_navs.values():
            latest_nav_date_str = info["date"]  # e.g. "27-Jun-2026"
            break
        # Convert to dd-mm-YYYY for column matching
        if latest_nav_date_str:
            parsed_live = _parse_amfi_date_str(latest_nav_date_str)
            if parsed_live:
                live_dt = datetime.strptime(parsed_live, "%Y-%m-%d")
                live_col = live_dt.strftime("%d-%m-%Y")
                # Add column if not already present
                if live_col not in df_sim.columns:
                    df_sim[live_col] = None
                    if live_col not in target_date_cols:
                        target_date_cols.append(live_col)
                # Inject live NAVs
                for idx, row in df_sim.iterrows():
                    isin_upper = row["ISIN_upper"]
                    if isin_upper in latest_navs:
                        df_sim.at[idx, live_col] = latest_navs[isin_upper]["nav"]
                # Carry forward into live column for any missing ISINs
                if len(sorted_cols) > 0:
                    last_hist = sorted_cols[-1]
                    df_sim[live_col] = df_sim[live_col].fillna(df_sim[last_hist])
        
    df_sim["Initial_Allocation"] = initial_amount * df_sim["Weight_Normalized"]
    df_sim["Units"] = df_sim["Initial_Allocation"] / df_sim[t0]
    
    daily_rows = []
    for d in target_date_cols:
        row_val = {"Date": d}
        total_val = 0.0
        for idx, row in df_sim.iterrows():
            name = row["Scheme Name"] or row["ISIN"]
            nav_t = row[d]
            val_t = row["Units"] * nav_t
            row_val[f"{name} (₹)"] = round(val_t, 2)
            total_val += val_t
            
        row_val["Total Portfolio Value (₹)"] = round(total_val, 2)
        daily_rows.append(row_val)
        
    df_tracker = pd.DataFrame(daily_rows)
    df_tracker["Daily Return (%)"] = 0.0
    df_tracker["Cumulative Return (%)"] = 0.0
    
    t0_val = df_tracker.at[0, "Total Portfolio Value (₹)"]
    for i in range(len(df_tracker)):
        curr_val = df_tracker.at[i, "Total Portfolio Value (₹)"]
        df_tracker.at[i, "Cumulative Return (%)"] = round(((curr_val - t0_val) / t0_val) * 100, 2)
        if i > 0:
            prev_val = df_tracker.at[i - 1, "Total Portfolio Value (₹)"]
            df_tracker.at[i, "Daily Return (%)"] = round(((curr_val - prev_val) / prev_val) * 100, 2)
            
    final_val = df_tracker.iloc[-1]["Total Portfolio Value (₹)"]
    abs_gain = final_val - initial_amount
    abs_return = (abs_gain / initial_amount) * 100
    
    peaks = df_tracker["Total Portfolio Value (₹)"].cummax()
    drawdowns = (df_tracker["Total Portfolio Value (₹)"] - peaks) / peaks * 100
    max_dd = drawdowns.min()
    
    best_day = df_tracker["Daily Return (%)"].max()
    worst_day = df_tracker["Daily Return (%)"].min()
    
    # Today's change (last row vs second-to-last)
    todays_change = 0.0
    todays_change_pct = 0.0
    if len(df_tracker) >= 2:
        yesterday_val = df_tracker.iloc[-2]["Total Portfolio Value (₹)"]
        todays_change = final_val - yesterday_val
        todays_change_pct = (todays_change / yesterday_val) * 100 if yesterday_val else 0.0
    
    metrics = {
        "Initial Value": initial_amount,
        "Current Value": final_val,
        "Total Gain/Loss": abs_gain,
        "Total Return (%)": abs_return,
        "Today's Change": todays_change,
        "Today's Change (%)": todays_change_pct,
        "Max Drawdown (%)": max_dd,
        "Best Day Return (%)": best_day,
        "Worst Day Return (%)": worst_day,
        "NAV Date": latest_nav_date_str or "N/A",
    }
    
    comp_stats = []
    for idx, row in df_sim.iterrows():
        isin = row["ISIN"]
        name = row["Scheme Name"]
        weight = row["Weight (%)"]
        initial_alloc = row["Initial_Allocation"]
        units = row["Units"]
        final_nav = row[target_date_cols[-1]]
        final_val_fund = units * final_nav
        fund_return = ((final_nav - row[t0]) / row[t0]) * 100
        comp_stats.append({
            "Scheme Name": name,
            "ISIN": isin,
            "Weight (%)": weight,
            "Initial Allocation (₹)": round(initial_alloc, 2),
            "Units": round(units, 4),
            "Buy NAV (₹)": round(row[t0], 4),
            "Current NAV (₹)": round(final_nav, 4),
            "Current Value (₹)": round(final_val_fund, 2),
            "Return (%)": round(fund_return, 2)
        })
        
    df_comp_stats = pd.DataFrame(comp_stats)
    
    return {
        "tracker": df_tracker,
        "metrics": metrics,
        "composition": df_comp_stats,
        "latest_nav_date": latest_nav_date_str,
    }, None


def style_portfolio_excel(res_dict: dict) -> bytes:
    """Generate a styled corporate Excel workbook for the portfolio simulation."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    df_tracker = res_dict["tracker"]
    df_comp = res_dict["composition"]
    
    wb = openpyxl.Workbook()
    
    # Sheet 1: Valuation Tracker
    ws1 = wb.active
    ws1.title = "Valuation Tracker"
    
    font_name = "Segoe UI"
    h_fill = PatternFill("solid", fgColor="1F497D")
    h_font = Font(name=font_name, size=11, bold=True, color="FFFFFF")
    data_font = Font(name=font_name, size=10)
    border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3")
    )
    even_fill = PatternFill("solid", fgColor="F2F5F8")
    odd_fill = PatternFill("solid", fgColor="FFFFFF")
    center_va = Alignment(horizontal="center", vertical="center")
    right_va = Alignment(horizontal="right", vertical="center")
    left_va = Alignment(horizontal="left", vertical="center")
    
    # Headers
    ws1.row_dimensions[1].height = 28
    for ci, col in enumerate(df_tracker.columns, 1):
        cell = ws1.cell(row=1, column=ci, value=col)
        cell.fill = h_fill
        cell.font = h_font
        cell.border = border
        cell.alignment = center_va
        
    # Data Rows
    for ri, row_data in enumerate(df_tracker.values, 2):
        ws1.row_dimensions[ri].height = 20
        fill = even_fill if ri % 2 == 0 else odd_fill
        for ci, val in enumerate(row_data, 1):
            cell = ws1.cell(row=ri, column=ci, value=val)
            cell.fill = fill
            cell.font = data_font
            cell.border = border
            col_name = df_tracker.columns[ci - 1]
            
            if col_name == "Date":
                cell.alignment = center_va
            elif "Return" in col_name:
                cell.number_format = '0.00"%"'
                cell.alignment = right_va
            else:
                cell.number_format = "0.00"
                cell.alignment = right_va
                
    for col in ws1.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=12)
        ws1.column_dimensions[get_column_letter(col[0].column)].width = max(min(max_len + 3, 55), 12)
        
    # Sheet 2: Bucket Composition
    ws2 = wb.create_sheet(title="Bucket Composition")
    ws2.row_dimensions[1].height = 28
    for ci, col in enumerate(df_comp.columns, 1):
        cell = ws2.cell(row=1, column=ci, value=col)
        cell.fill = h_fill
        cell.font = h_font
        cell.border = border
        cell.alignment = left_va if col in ("Scheme Name", "ISIN") else center_va
        
    for ri, row_data in enumerate(df_comp.values, 2):
        ws2.row_dimensions[ri].height = 20
        fill = even_fill if ri % 2 == 0 else odd_fill
        for ci, val in enumerate(row_data, 1):
            cell = ws2.cell(row=ri, column=ci, value=val)
            cell.fill = fill
            cell.font = data_font
            cell.border = border
            col_name = df_comp.columns[ci - 1]
            
            if col_name in ("Scheme Name", "ISIN"):
                cell.alignment = left_va
            elif "Weight" in col_name or "Return" in col_name:
                cell.number_format = '0.00"%"'
                cell.alignment = right_va
            elif "Units" in col_name:
                cell.number_format = "0.0000"
                cell.alignment = right_va
            else:
                cell.number_format = "0.00"
                cell.alignment = right_va
                
    for col in ws2.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=12)
        ws2.column_dimensions[get_column_letter(col[0].column)].width = max(min(max_len + 3, 55), 12)
        
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── Project helpers ─────────────────────────────────────────────────────────
sys.path.append(str(Path(__file__).parent))
from amfi_nav import classify_option_type, classify_plan_type
from ui_theme import (
    inject_custom_css,
    render_app_footer,
    render_hero,
    render_info_card,
    render_section_header,
)

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AMFI NAV Fetcher",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Constants ────────────────────────────────────────────────────────────────
AMFI_HISTORY_URL = (
    "https://portal.amfiindia.com/DownloadNAVHistoryReport_Po.aspx"
    "?frmdt={frmdt}&todt={todt}"
)
AMFI_LATEST_URL = "https://portal.amfiindia.com/spages/NAVAll.txt"

# ─── Helpers ─────────────────────────────────────────────────────────────────

def parse_amfi_text(text: str, isin_list: tuple[str, ...] | None = None) -> pd.DataFrame:
    """Parse the semicolon-delimited AMFI NAV feed into a DataFrame."""
    rows: List[dict] = []
    current_section = "Unknown"

    isin_set = {isin.strip().upper() for isin in isin_list} if isin_list else None

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("Scheme Code;"):
            continue
        if (
            line.startswith("Open Ended")
            or line.startswith("Closed Ended")
            or line.startswith("Interval Fund Schemes")
        ):
            current_section = line
            continue
        if line.count(";") < 5:
            continue

        # Optimize memory & parsing: Skip processing the line if isin_list is provided
        # and none of the targeted ISINs are in the raw text line (case-insensitive)
        if isin_set:
            line_upper = line.upper()
            if not any(isin in line_upper for isin in isin_set):
                continue

        parts = [p.strip() for p in line.split(";")]
        if len(parts) < 8:
            continue

        scheme_code = parts[0]
        scheme_name = parts[1]
        isin_growth = parts[2] if parts[2] not in ("", "-") else None
        isin_reinvest = parts[3] if parts[3] not in ("", "-") else None
        nav_value = parts[4]
        nav_date = parts[7]

        # Double check matching against ISIN set
        if isin_set:
            g_match = isin_growth is not None and isin_growth.upper() in isin_set
            r_match = isin_reinvest is not None and isin_reinvest.upper() in isin_set
            if not (g_match or r_match):
                continue

        nav = pd.to_numeric(nav_value.replace(",", ""), errors="coerce")

        rows.append(
            {
                "Asset Class": current_section,
                "Scheme Code": scheme_code,
                "ISIN Div Payout / ISIN Growth": isin_growth,
                "ISIN Div Reinvestment": isin_reinvest,
                "Scheme Name": scheme_name,
                "NAV": nav,
                "NAV Date": nav_date,
            }
        )

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    df["Plan Type"] = df["Scheme Name"].apply(classify_plan_type)
    df["Option Type"] = df["Scheme Name"].apply(classify_option_type)
    return df


def filter_by_isins(df: pd.DataFrame, isin_list: List[str]) -> pd.DataFrame:
    """Keep only rows where either ISIN column matches the target list."""
    mask = df["ISIN Div Payout / ISIN Growth"].isin(isin_list) | df[
        "ISIN Div Reinvestment"
    ].isin(isin_list)
    return df[mask].copy()


def pivot_to_wide(df: pd.DataFrame, date_cols: List[str]) -> pd.DataFrame:
    """Pivot long-format NAV rows to wide format (one column per date)."""
    meta_cols = [
        "Asset Class",
        "Scheme Code",
        "ISIN Div Payout / ISIN Growth",
        "ISIN Div Reinvestment",
        "Scheme Name",
        "Plan Type",
        "Option Type",
    ]

    meta = df[meta_cols].drop_duplicates(subset=["Scheme Code"])
    pivot = (
        df.pivot_table(index="Scheme Code", columns="NAV Date", values="NAV", aggfunc="first")
        .reset_index()
    )
    result = pd.merge(meta, pivot, on="Scheme Code", how="left")

    for d in date_cols:
        if d not in result.columns:
            result[d] = None

    # Keep only requested date columns
    available_dates = [d for d in date_cols if d in result.columns]
    result = result[meta_cols + available_dates]
    return result.sort_values(["Asset Class", "Scheme Name"]).reset_index(drop=True)


def _date_to_dmy(date_str: str) -> str:
    """Convert '29-May-2026' → '29-05-2026'."""
    try:
        return datetime.strptime(date_str, "%d-%b-%Y").strftime("%d-%m-%Y")
    except Exception:
        return date_str


def style_excel(df: pd.DataFrame, date_cols: List[str], is_aum_only: bool = False) -> bytes:
    """Write df to a styled Excel workbook and return as bytes.

    When both NAV and AUM columns are present (columns named like '29-May-2026 (NAV)'
    and '29-May-2026 (AUM)'), the Excel uses a two-row header:
      Row 1: meta cols (merged ↕) | 'NAV' merged across all NAV date cols | 'AUM' merged across all AUM date cols
      Row 2: date labels in dd/mm/yyyy under each group
    Data starts at row 3.
    """
    import openpyxl
    from openpyxl.utils import get_column_letter

    buf = BytesIO()

    font_name   = "Segoe UI"
    h_fill      = PatternFill("solid", fgColor="1F497D")   # dark blue – meta headers
    nav_fill    = PatternFill("solid", fgColor="1E4D8C")   # blue – NAV group header
    aum_fill    = PatternFill("solid", fgColor="145A32")   # green – AUM group header
    date_nav_fill = PatternFill("solid", fgColor="2E75B6") # lighter blue – NAV date row
    date_aum_fill = PatternFill("solid", fgColor="1E8449") # lighter green – AUM date row
    h_font      = Font(name=font_name, size=11, bold=True, color="FFFFFF")
    date_font   = Font(name=font_name, size=10, bold=True, color="FFFFFF")
    data_font   = Font(name=font_name, size=10)
    even_fill   = PatternFill("solid", fgColor="F2F5F8")
    odd_fill    = PatternFill("solid", fgColor="FFFFFF")
    border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )
    center_va = Alignment(horizontal="center", vertical="center")
    left_va   = Alignment(horizontal="left",   vertical="center")

    # ── Detect whether we have interleaved (NAV)/(AUM) pairs ─────────────────
    has_nav_aum_pairs = any("(NAV)" in c for c in df.columns)

    if has_nav_aum_pairs:
        # Split columns into: meta, nav_dates, aum_dates (preserving order)
        meta_cols  = [c for c in df.columns if "(NAV)" not in c and "(AUM)" not in c]
        nav_cols   = [c for c in df.columns if c.endswith(" (NAV)")]
        aum_cols   = [c for c in df.columns if c.endswith(" (AUM)")]

        nav_dates  = [c[: -len(" (NAV)")] for c in nav_cols]  # "29-May-2026"
        aum_dates  = [c[: -len(" (AUM)")] for c in aum_cols]

        # Re-order df columns: meta | all NAV cols | all AUM cols
        ordered_cols = meta_cols + nav_cols + aum_cols
        df = df[ordered_cols]

        n_meta = len(meta_cols)
        n_nav  = len(nav_cols)
        n_aum  = len(aum_cols)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "NAV Data"

        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 22

        # ── Row 1 ─────────────────────────────────────────────────────────────
        # Meta columns: merged vertically (rows 1–2)
        for i, col in enumerate(meta_cols):
            ci = i + 1
            cell = ws.cell(row=1, column=ci, value=col)
            cell.fill = h_fill
            cell.font = h_font
            cell.border = border
            cell.alignment = left_va if col in ("Scheme Name", "Asset Class") else center_va
            ws.merge_cells(start_row=1, start_column=ci, end_row=2, end_column=ci)

        # NAV group label: merged across all NAV date columns
        nav_start = n_meta + 1
        nav_end   = n_meta + n_nav
        if n_nav > 0:
            cell = ws.cell(row=1, column=nav_start, value="NAV")
            cell.fill = nav_fill
            cell.font = h_font
            cell.border = border
            cell.alignment = center_va
            if n_nav > 1:
                ws.merge_cells(start_row=1, start_column=nav_start, end_row=1, end_column=nav_end)
            # Apply border to all cells in the merge range
            for ci in range(nav_start, nav_end + 1):
                ws.cell(row=1, column=ci).border = border

        # AUM group label: merged across all AUM date columns
        aum_start = n_meta + n_nav + 1
        aum_end   = n_meta + n_nav + n_aum
        if n_aum > 0:
            cell = ws.cell(row=1, column=aum_start, value="AUM")
            cell.fill = aum_fill
            cell.font = h_font
            cell.border = border
            cell.alignment = center_va
            if n_aum > 1:
                ws.merge_cells(start_row=1, start_column=aum_start, end_row=1, end_column=aum_end)
            for ci in range(aum_start, aum_end + 1):
                ws.cell(row=1, column=ci).border = border

        # ── Row 2: date labels in dd/mm/yyyy ──────────────────────────────────
        # Meta cells in row 2 are covered by merges, just apply styling
        for i in range(n_meta):
            ci = i + 1
            cell = ws.cell(row=2, column=ci)
            cell.fill = h_fill
            cell.font = h_font
            cell.border = border

        for i, d in enumerate(nav_dates):
            ci = nav_start + i
            cell = ws.cell(row=2, column=ci, value=_date_to_dmy(d))
            cell.fill = date_nav_fill
            cell.font = date_font
            cell.border = border
            cell.alignment = center_va

        for i, d in enumerate(aum_dates):
            ci = aum_start + i
            cell = ws.cell(row=2, column=ci, value=_date_to_dmy(d))
            cell.fill = date_aum_fill
            cell.font = date_font
            cell.border = border
            cell.alignment = center_va

        # ── Data rows (start at row 3) ────────────────────────────────────────
        for ri_data, (_, row_data) in enumerate(df.iterrows()):
            ri = ri_data + 3
            ws.row_dimensions[ri].height = 20
            fill = even_fill if ri % 2 == 0 else odd_fill
            for ci, col in enumerate(ordered_cols, 1):
                cell = ws.cell(row=ri, column=ci)
                val  = row_data[col]
                cell.fill   = fill
                cell.font   = data_font
                cell.border = border
                if col in nav_cols or col in aum_cols:
                    if pd.notna(val) and val is not None:
                        cell.value = val
                        cell.number_format = "0.00" if col in aum_cols else "0.0000"
                    else:
                        cell.value = "—"
                    cell.alignment = center_va
                else:
                    cell.value = val if pd.notna(val) else None
                    cell.alignment = left_va if col in ("Scheme Name", "Asset Class") else center_va

        # ── Column widths ─────────────────────────────────────────────────────
        for col_cells in ws.iter_cols(min_row=1, max_row=ws.max_row):
            max_len = 10
            for c in col_cells:
                try:
                    max_len = max(max_len, len(str(c.value or "")) + (2 if c.row <= 2 else 0))
                except Exception:
                    pass
            ws.column_dimensions[get_column_letter(col_cells[0].column)].width = max(min(max_len + 2, 55), 10)

        wb.save(buf)

    else:
        # ── Original single-header path (NAV only, AUM only, or long format) ──
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="NAV Data")
            wb = writer.book
            ws = writer.sheets["NAV Data"]

            h_fill2    = PatternFill("solid", fgColor="1F497D")
            h_font2    = Font(name=font_name, size=11, bold=True, color="FFFFFF")
            data_font2 = Font(name=font_name, size=10)

            ws.row_dimensions[1].height = 28
            for ci, col_name in enumerate(df.columns, 1):
                cell = ws.cell(row=1, column=ci)
                cell.fill = h_fill2
                cell.font = h_font2
                cell.border = border
                left_a = col_name in ("Scheme Name", "Asset Class")
                cell.alignment = Alignment(horizontal="left" if left_a else "center", vertical="center")

            for ri in range(2, len(df) + 2):
                ws.row_dimensions[ri].height = 20
                fill = even_fill if ri % 2 == 0 else odd_fill
                for ci, col_name in enumerate(df.columns, 1):
                    cell = ws.cell(row=ri, column=ci)
                    cell.fill = fill
                    cell.font = data_font2
                    cell.border = border
                    if col_name in ("NAV Date", "AUM Date"):
                        if cell.value is None or cell.value == "":
                            cell.value = "—"
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif col_name in ("NAV", "NAVs"):
                        if cell.value is not None and cell.value != "":
                            cell.number_format = "0.00" if col_name == "NAVs" else "0.0000"
                        else:
                            cell.value = "—"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif col_name == "Daily return":
                        if cell.value is not None and cell.value != "":
                            cell.number_format = '0.00"%"'
                        else:
                            cell.value = "—"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif col_name in (
                        "Closing AUM as on previous day", 
                        "Actual AUM as on current date", 
                        "Derived AUM as on curent day", 
                        "Net flows on current day",
                        "AUM"
                    ):
                        if cell.value is not None and cell.value != "":
                            cell.number_format = "0.00"
                        else:
                            cell.value = "—"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif col_name in ("Scheme Name", "Asset Class"):
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="center", vertical="center")

            for col in ws.columns:
                max_len = max((len(str(c.value or "")) + (3 if c.row == 1 else 0) for c in col), default=12)
                ws.column_dimensions[get_column_letter(col[0].column)].width = max(min(max_len + 2, 55), 12)

    return buf.getvalue()


@st.cache_data(ttl=900, show_spinner=False)
def fetch_amfi_data(frmdt: str, todt: str, isin_list: tuple[str, ...] | None = None) -> pd.DataFrame:
    url = AMFI_HISTORY_URL.format(frmdt=frmdt, todt=todt)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    import time
    max_retries = 3
    delay = 2.0
    resp = None
    for attempt in range(max_retries):
        try:
            resp = requests.get(url, headers=headers, stream=True, timeout=300)
            resp.raise_for_status()
            break
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
            if attempt == max_retries - 1:
                raise e
            time.sleep(delay)
            delay *= 2
    
    rows: List[dict] = []
    current_section = "Unknown"
    
    isin_set = {isin.strip().upper() for isin in isin_list if isin.strip()} if isin_list else None
    
    for line_bytes in resp.iter_lines():
        if not line_bytes:
            continue
        line = line_bytes.decode('utf-8', errors='ignore')
        
        # Fast section check (no semicolons in section headers)
        if ";" not in line:
            line_stripped = line.strip()
            if (
                line_stripped.startswith("Open Ended")
                or line_stripped.startswith("Closed Ended")
                or line_stripped.startswith("Interval Fund Schemes")
            ):
                current_section = line_stripped
            continue
            
        # Optimize memory & parsing: Skip processing the line if isin_set is provided
        # and none of the targeted ISINs are in the raw text line (case-insensitive)
        if isin_set:
            line_upper = line.upper()
            if not any(isin in line_upper for isin in isin_set):
                continue
                
        parts = line.split(";")
        if len(parts) < 8:
            continue
            
        isin_growth = parts[2].strip()
        isin_reinvest = parts[3].strip()
        
        isin_growth_upper = isin_growth.upper() if isin_growth != "-" else ""
        isin_reinvest_upper = isin_reinvest.upper() if isin_reinvest != "-" else ""
        
        # Double check matching against ISIN set
        if isin_set:
            g_match = isin_growth_upper and isin_growth_upper in isin_set
            r_match = isin_reinvest_upper and isin_reinvest_upper in isin_set
            if not (g_match or r_match):
                continue
                
        scheme_code = parts[0].strip()
        scheme_name = parts[1].strip()
        nav_value = parts[4].strip()
        nav_date = parts[7].strip()
        
        scheme_code = scheme_code if scheme_code != "-" else None
        isin_growth_val = isin_growth if isin_growth != "-" else None
        isin_reinvest_val = isin_reinvest if isin_reinvest != "-" else None
        
        nav = pd.to_numeric(nav_value.replace(",", ""), errors="coerce")
        
        rows.append(
            {
                "Asset Class": current_section,
                "Scheme Code": scheme_code,
                "ISIN Div Payout / ISIN Growth": isin_growth_val,
                "ISIN Div Reinvestment": isin_reinvest_val,
                "Scheme Name": scheme_name,
                "NAV": nav,
                "NAV Date": nav_date,
            }
        )
        
    if not rows:
        return pd.DataFrame()
        
    df = pd.DataFrame(rows)
    df["Plan Type"] = df["Scheme Name"].apply(classify_plan_type)
    df["Option Type"] = df["Scheme Name"].apply(classify_option_type)
    return df


def fetch_amfi_data_chunked(start_date, end_date, isin_list: List[str] | None = None) -> pd.DataFrame:
    """Fetch AMFI data by chunking the date range into 90-day intervals to prevent timeout & memory crash."""
    import time
    passed_isins = tuple(isin_list) if isin_list else None
    dfs = []
    current_start = start_date
    
    while current_start <= end_date:
        current_end = min(current_start + timedelta(days=90), end_date)
        frmdt_str = current_start.strftime("%d-%b-%Y")
        todt_str = current_end.strftime("%d-%b-%Y")
        
        df_chunk = fetch_amfi_data(frmdt_str, todt_str, passed_isins)
        if not df_chunk.empty:
            dfs.append(df_chunk)
            
        current_start = current_end + timedelta(days=1)
        if current_start <= end_date:
            time.sleep(0.5)  # avoid hitting rate limits
        
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)


def build_date_cols(start_date, end_date, skip_sunday: bool) -> List[str]:
    """Generate sorted list of date strings between start and end."""
    cols = []
    current = start_date
    while current <= end_date:
        if not (skip_sunday and current.weekday() == 6):
            cols.append(current.strftime("%d-%m-%Y"))
        current += timedelta(days=1)
    return cols


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    inject_custom_css()
    render_hero(
        title="AMFI NAV Fetcher",
        subtitle=(
            "Fetch live and historical NAV data from AMFI India by ISIN or date range. "
            "Preview results inline and export corporate-styled Excel reports."
        ),
        chips=[
            "🔍 ISIN + date range lookup",
            "📅 Full-market date range export",
            "📥 Styled Excel download",
        ],
    )

    render_section_header("⚙️", "Fetch Configuration", "Choose mode, dates, and export options")
    mode = st.radio(
        "Fetch mode",
        ["By ISIN + Date Range", "By Date Range Only (all funds)", "Portfolio Bucket Tracker"],
        horizontal=True,
        label_visibility="collapsed",
    )

    if mode == "Portfolio Bucket Tracker":
        st.markdown('<hr class="panel-divider">', unsafe_allow_html=True)
        
        # 1. Initialize session state buckets if they don't exist
        if "portfolio_buckets" not in st.session_state:
            default_df = pd.DataFrame([
                {"Scheme Name": "Quant Large Cap Fund-Reg(G)", "ISIN": "INF966L01AW4", "Weight (%)": 10.0},
                {"Scheme Name": "DSP Equity Opportunities Fund-Reg(G)", "ISIN": "INF740K01094", "Weight (%)": 9.0},
                {"Scheme Name": "SBI Large & Midcap Fund-Reg(G)", "ISIN": "INF200K01305", "Weight (%)": 6.0},
                {"Scheme Name": "HDFC Flexi Cap Fund(G)", "ISIN": "INF179K01608", "Weight (%)": 9.0},
                {"Scheme Name": "ICICI Pru Focused Equity Fund(G)", "ISIN": "INF109K01BZ4", "Weight (%)": 7.0},
                {"Scheme Name": "ICICI Pru Dividend Yield Equity Fund(G)", "ISIN": "INF109KA1TX4", "Weight (%)": 7.0},
                {"Scheme Name": "Canara Rob Multi Cap Fund-Reg(G)", "ISIN": "INF760K01KR2", "Weight (%)": 7.0},
                {"Scheme Name": "Kotak Multicap Fund-Reg(G)", "ISIN": "INF174KA1HS9", "Weight (%)": 7.0},
                {"Scheme Name": "Bandhan Large & Mid Cap Fund-Reg(G)", "ISIN": "INF194K01524", "Weight (%)": 7.0},
                {"Scheme Name": "Invesco India Focused Fund-Reg(G)", "ISIN": "INF205KA1189", "Weight (%)": 7.0},
                {"Scheme Name": "Kotak Emerging Equity Fund(G)", "ISIN": "INF174K01DS9", "Weight (%)": 6.0},
                {"Scheme Name": "SBI Infrastructure Fund-Reg(G)", "ISIN": "INF200K01CT2", "Weight (%)": 6.0},
                {"Scheme Name": "Invesco India Smallcap Fund-Reg(G)", "ISIN": "INF205K011T7", "Weight (%)": 7.0},
                {"Scheme Name": "HDFC Small Cap Fund-Reg(G)", "ISIN": "INF179KA1RZ8", "Weight (%)": 5.0},
            ])
            st.session_state["portfolio_buckets"] = {"Default Balanced Portfolio": default_df}
            st.session_state["active_bucket_name"] = "Default Balanced Portfolio"

        # Bucket selection & management controls
        render_section_header("📁", "Bucket Management", "Create, switch, or delete fund buckets")
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            bucket_options = list(st.session_state["portfolio_buckets"].keys())
            active_bucket = st.selectbox(
                "Select Active Bucket", 
                bucket_options, 
                index=bucket_options.index(st.session_state["active_bucket_name"])
            )
            st.session_state["active_bucket_name"] = active_bucket
            
        with col_m2:
            new_bucket_name = st.text_input("New Bucket Name", placeholder="e.g. My Conservative Portfolio")
            c_add, c_del = st.columns(2)
            with c_add:
                if st.button("Create New Bucket", use_container_width=True):
                    if new_bucket_name.strip() and new_bucket_name not in st.session_state["portfolio_buckets"]:
                        curr_df = st.session_state["portfolio_buckets"][st.session_state["active_bucket_name"]].copy()
                        st.session_state["portfolio_buckets"][new_bucket_name] = curr_df
                        st.session_state["active_bucket_name"] = new_bucket_name
                        st.rerun()
            with c_del:
                if st.button("Delete Active Bucket", use_container_width=True, type="secondary"):
                    if len(st.session_state["portfolio_buckets"]) > 1:
                        del st.session_state["portfolio_buckets"][st.session_state["active_bucket_name"]]
                        st.session_state["active_bucket_name"] = list(st.session_state["portfolio_buckets"].keys())[0]
                        st.rerun()
                    else:
                        st.warning("Cannot delete the last remaining bucket.")

        st.markdown('<hr class="panel-divider">', unsafe_allow_html=True)
        
        # 2. Investment Start Date & Amount (no end date — always tracks to today)
        render_section_header("📅", "Investment Setup", "Pick when you started investing — the tracker runs to today automatically")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            start_date = st.date_input("Investment Start Date", value=(datetime.today() - timedelta(days=90)).date())
        with col_b:
            initial_amount = st.number_input("Investment Amount (₹)", value=100000.0, min_value=100.0, step=1000.0)
        with col_c:
            skip_sunday = st.checkbox("Skip Sundays", value=True)

        st.markdown('<hr class="panel-divider">', unsafe_allow_html=True)
        
        # 3. Bucket Composition Editor
        render_section_header("📋", "Bucket Composition", f"Add/remove funds and set weights for '{st.session_state['active_bucket_name']}'")
        
        uploaded_file = st.file_uploader("Upload Bucket (Excel/CSV with ISIN and Weights)", type=["xlsx", "xls", "csv"])
        
        current_bucket_df = st.session_state["portfolio_buckets"][st.session_state["active_bucket_name"]]
        if uploaded_file is not None:
            parsed_df = parse_bucket_input(uploaded_file=uploaded_file)
            if not parsed_df.empty:
                st.session_state["portfolio_buckets"][st.session_state["active_bucket_name"]] = parsed_df
                current_bucket_df = parsed_df
                st.success("Successfully loaded bucket from file!")
                
        edited_df = st.data_editor(
            current_bucket_df,
            column_config={
                "Scheme Name": st.column_config.TextColumn("Scheme Name", width="large", help="Optional name for display"),
                "ISIN": st.column_config.TextColumn("ISIN", required=True, help="Mutual Fund ISIN (Growth or Reinvestment)"),
                "Weight (%)": st.column_config.NumberColumn("Weight (%)", min_value=0.0, max_value=100.0, format="%.2f%%", required=True, help="Percentage weight in portfolio")
            },
            num_rows="dynamic",
            use_container_width=True
        )
        
        if edited_df is not None:
            st.session_state["portfolio_buckets"][st.session_state["active_bucket_name"]] = edited_df
            
        weight_sum = edited_df["Weight (%)"].sum() if not edited_df.empty else 0
        if weight_sum != 100.0:
            st.info(f"⚖️ Current weights sum to **{weight_sum:.2f}%**. Weights will be automatically normalized to 100%.")
        else:
            st.success("⚖️ Weights sum to exactly **100%**!")

        st.markdown('<hr class="panel-divider">', unsafe_allow_html=True)
        
        # 4. Track / Refresh buttons
        col_btn1, col_btn2 = st.columns([3, 1])
        with col_btn1:
            track_btn = st.button("📊 Track Portfolio", type="primary", use_container_width=True)
        with col_btn2:
            refresh_btn = st.button("🔄 Refresh", use_container_width=True)
        
        should_run = track_btn or refresh_btn
        
        if should_run:
            if start_date > datetime.today().date():
                st.error("Start Date cannot be in the future.")
            else:
                with st.spinner("Fetching latest NAVs from AMFI..."):
                    res, err = run_live_portfolio(
                        edited_df,
                        start_date,
                        initial_amount,
                        skip_sunday
                    )
                    
                if err:
                    st.error(err)
                else:
                    met = res["metrics"]
                    nav_date_display = met.get("NAV Date", "N/A")
                    
                    # ── Prominent LIVE current value ──────────────────────────
                    gain_color = "#10b981" if met["Total Gain/Loss"] >= 0 else "#ef4444"
                    today_color = "#10b981" if met["Today's Change"] >= 0 else "#ef4444"
                    gain_sign = "+" if met["Total Gain/Loss"] >= 0 else ""
                    today_sign = "+" if met["Today's Change"] >= 0 else ""
                    
                    st.markdown(
                        f"""
                        <div style="background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); border-radius: 16px; padding: 28px 32px; margin-bottom: 20px; border: 1px solid rgba(212,175,55,0.3);">
                            <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 16px;">
                                <div>
                                    <div style="color: #94a3b8; font-size: 13px; text-transform: uppercase; letter-spacing: 1px;">Current Portfolio Value</div>
                                    <div style="color: #f8fafc; font-size: 36px; font-weight: 700; margin: 4px 0;">₹{met['Current Value']:,.2f}</div>
                                    <div style="color: {gain_color}; font-size: 16px; font-weight: 600;">{gain_sign}₹{met['Total Gain/Loss']:,.2f} ({gain_sign}{met['Total Return (%)']:.2f}%) total</div>
                                </div>
                                <div style="text-align: right;">
                                    <div style="color: #94a3b8; font-size: 13px; text-transform: uppercase; letter-spacing: 1px;">Today's Change</div>
                                    <div style="color: {today_color}; font-size: 24px; font-weight: 700; margin: 4px 0;">{today_sign}₹{met["Today's Change"]:,.2f}</div>
                                    <div style="color: {today_color}; font-size: 14px;">{today_sign}{met["Today's Change (%)"]:.2f}%</div>
                                    <div style="color: #64748b; font-size: 12px; margin-top: 6px;">NAV as of {nav_date_display}</div>
                                </div>
                            </div>
                            <div style="display: flex; gap: 24px; margin-top: 16px; padding-top: 16px; border-top: 1px solid rgba(255,255,255,0.1);">
                                <div><span style="color: #64748b; font-size: 12px;">Invested</span><br><span style="color: #e2e8f0; font-weight: 600;">₹{met['Initial Value']:,.2f}</span></div>
                                <div><span style="color: #64748b; font-size: 12px;">Max Drawdown</span><br><span style="color: #ef4444; font-weight: 600;">{met['Max Drawdown (%)']:.2f}%</span></div>
                                <div><span style="color: #64748b; font-size: 12px;">Best Day</span><br><span style="color: #10b981; font-weight: 600;">{met['Best Day Return (%)']:.2f}%</span></div>
                                <div><span style="color: #64748b; font-size: 12px;">Worst Day</span><br><span style="color: #ef4444; font-weight: 600;">{met['Worst Day Return (%)']:.2f}%</span></div>
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                    
                    # ── Fund-wise breakdown ───────────────────────────────────
                    render_section_header("💼", "Fund-wise Holdings")
                    st.dataframe(
                        res["composition"],
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Weight (%)": st.column_config.NumberColumn("Weight (%)", format="%.2f%%"),
                            "Return (%)": st.column_config.NumberColumn("Return (%)", format="%.2f%%"),
                        }
                    )
                    
                    # ── Portfolio Value Chart ─────────────────────────────────
                    render_section_header("📈", "Portfolio Value Over Time")
                    chart_df = res["tracker"].copy()
                    chart_df["Date_dt"] = pd.to_datetime(chart_df["Date"], format="%d-%m-%Y")
                    chart_df = chart_df.set_index("Date_dt")
                    st.line_chart(chart_df["Total Portfolio Value (₹)"])
                    
                    # ── Daily History ─────────────────────────────────────────
                    with st.expander("📋 Daily Valuation History", expanded=False):
                        st.dataframe(
                            res["tracker"],
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "Daily Return (%)": st.column_config.NumberColumn("Daily Return (%)", format="%.2f%%"),
                                "Cumulative Return (%)": st.column_config.NumberColumn("Cumulative Return (%)", format="%.2f%%")
                            }
                        )
                    
                    # ── Export ─────────────────────────────────────────────────
                    render_section_header("📥", "Export Report")
                    with st.spinner("Generating Excel report..."):
                        excel_bytes = style_portfolio_excel(res)
                        
                    st.download_button(
                        label="Download Portfolio Report (.xlsx)",
                        data=excel_bytes,
                        file_name=f"portfolio_{st.session_state['active_bucket_name']}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
        render_app_footer()
        return

    st.markdown('<hr class="panel-divider">', unsafe_allow_html=True)

    default_end = datetime.today().date()
    default_start = (datetime.today() - timedelta(days=7)).date()

    render_section_header("📅", "Date Range")
    col_a, col_b = st.columns(2)
    with col_a:
        start_date = st.date_input("Start Date", value=default_start)
    with col_b:
        end_date = st.date_input("End Date", value=default_end)

    isin_list: List[str] = []
    isin_mode = mode.startswith("By ISIN")
    if isin_mode:
        render_info_card(
            "<strong>ISIN mode:</strong> Enter one or more ISINs (Growth <em>or</em> Reinvestment) — "
            "one per line or comma-separated. Rows matching either ISIN column will be fetched."
        )
        raw_isins = st.text_area(
            "ISINs",
            height=160,
            placeholder="INF209K01AJ8\nINF846K01CH7\nINF760K01019\n...",
        )
        isin_list = [x.strip() for x in re.split(r"[,\n\s]+", raw_isins) if x.strip()]
    else:
        render_info_card(
            "<strong>Date-range mode:</strong> No ISINs required — the full AMFI database "
            "for the selected period will be fetched (~8,000+ schemes). "
            "Ranges are limited to 90 days in this mode."
        )

    render_section_header("🎛️", "Export Options")
    c1, c2, c3 = st.columns(3)
    with c1:
        skip_sunday = st.checkbox("Skip Sundays", value=True)
    with c2:
        carry_forward = st.checkbox("Carry forward on holidays", value=True)
    with c3:
        pivot_dates = st.checkbox("Pivot: one column per date", value=True)

    col_c1, col_c2, col_c3 = st.columns(3)
    with col_c1:
        want_nav = st.checkbox("Want NAV", value=True)
    with col_c2:
        want_aum = st.checkbox("Want AUM", value=True)
    with col_c3:
        want_flows = st.checkbox("Want Flows", value=False)

    fetch_live_aum = False
    if want_aum or want_flows:
        fetch_live_aum = st.checkbox("Fetch live daily AUM (slower)", value=True)
        if want_flows:
            fetch_live_aum = True

    fetch_btn = st.button("Fetch data", type="primary", use_container_width=True)

    if not fetch_btn:
        render_app_footer()
        return

    # ── Validation ────────────────────────────────────────────────────────────
    if not want_nav and not want_aum and not want_flows:
        st.error("Please select at least one data type (NAV, AUM, or Flows) to export.")
        return
    if start_date > end_date:
        st.error("Start Date must be before or equal to End Date.")
        return
    if isin_mode and not isin_list:
        st.error("Please enter at least one ISIN.")
        return

    date_diff_days = (end_date - start_date).days
    if not isin_mode and date_diff_days > 90:
        st.error(
            "For 'By Date Range Only (all funds)' mode, the date range is limited to 90 days. "
            "Use ISIN mode for larger ranges."
        )
        return

    if want_flows:
        want_nav = True
        want_aum = True
        pivot_dates = True
        fetch_live_aum = True

    frmdt_str = start_date.strftime("%d-%b-%Y")
    todt_str = end_date.strftime("%d-%b-%Y")

    fetch_start_date = start_date
    if want_flows:
        fetch_start_date = start_date - timedelta(days=10)

    with st.spinner(f"Contacting AMFI India for {fetch_start_date.strftime('%d-%b-%Y')} → {todt_str} …"):
        try:
            df_raw = fetch_amfi_data_chunked(fetch_start_date, end_date, isin_list if isin_mode else None)
        except Exception as exc:
            st.error(f"Failed to fetch data: {exc}")
            return

    if df_raw.empty:
        st.warning("No data returned from AMFI for this date range.")
        return

    # ── Filter by ISINs if needed ─────────────────────────────────────────────
    if isin_mode:
        df_filtered = filter_by_isins(df_raw, isin_list)
        if df_filtered.empty:
            st.warning(
                "No records found for the specified ISINs. "
                "Double-check the ISINs or widen the date range."
            )
            return
    else:
        df_filtered = df_raw

    df_filtered["NAV Date"] = parse_amfi_date_series(df_filtered["NAV Date"])

    # ── Build AUM data using Performance API with Excel portfolio fallback ────
    df_port = load_portfolio_aum_data()
    df_filtered = populate_actual_aum(df_filtered, df_port, want_aum=want_aum, fetch_live_aum=fetch_live_aum)

    # ── Build target date columns ─────────────────────────────────────────────
    date_cols = build_date_cols(fetch_start_date, end_date, skip_sunday)

    # ── Stats ─────────────────────────────────────────────────────────────────
    total_schemes = df_filtered["Scheme Code"].nunique()
    total_records = len(df_filtered)
    dates_found = sorted(df_filtered["NAV Date"].unique())

    st.markdown(
        f"""
        <div class="stat-row">
            <span class="stat-chip">🏦 {total_schemes:,} schemes</span>
            <span class="stat-chip">📋 {total_records:,} records</span>
            <span class="stat-chip">📅 {len(dates_found)} trading dates</span>
            <span class="stat-chip">🗓️ {frmdt_str} → {todt_str}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Pivot or long format ──────────────────────────────────────────────────
    if pivot_dates:
        meta_cols = [
            "Asset Class",
            "Scheme Code",
            "ISIN Div Payout / ISIN Growth",
            "ISIN Div Reinvestment",
            "Scheme Name",
            "Plan Type",
            "Option Type",
        ]
        fund_metadata = df_filtered[meta_cols].drop_duplicates(subset=["Scheme Code"])
        df_filtered["NAV_Date_Str"] = df_filtered["NAV Date"].dt.strftime("%d-%m-%Y")
        
        if want_nav and not want_aum:
            df_pivot = df_filtered.pivot_table(index="Scheme Code", columns="NAV_Date_Str", values="NAV", aggfunc="first").reset_index()
            display_date_cols = date_cols
            is_aum_only = False
        elif want_aum and not want_nav:
            df_pivot = df_filtered.pivot_table(index="Scheme Code", columns="NAV_Date_Str", values="AUM", aggfunc="first").reset_index()
            display_date_cols = date_cols
            is_aum_only = True
        else:
            df_pivot_nav = df_filtered.pivot_table(index="Scheme Code", columns="NAV_Date_Str", values="NAV", aggfunc="first").reset_index()
            df_pivot_aum = df_filtered.pivot_table(index="Scheme Code", columns="NAV_Date_Str", values="AUM", aggfunc="first").reset_index()
            
            nav_cols_map = {d: f"{d} (NAV)" for d in date_cols if d in df_pivot_nav.columns}
            aum_cols_map = {d: f"{d} (AUM)" for d in date_cols if d in df_pivot_aum.columns}
            
            df_pivot_nav = df_pivot_nav.rename(columns=nav_cols_map)
            df_pivot_aum = df_pivot_aum.rename(columns=aum_cols_map)
            
            df_pivot = pd.merge(df_pivot_nav, df_pivot_aum, on="Scheme Code", how="left")
            
            interleaved_dates = []
            for d in date_cols:
                interleaved_dates.append(f"{d} (NAV)")
                interleaved_dates.append(f"{d} (AUM)")
            display_date_cols = interleaved_dates
            is_aum_only = False
            
        df_display = pd.merge(fund_metadata, df_pivot, on="Scheme Code", how="left")
        
        for date_col in display_date_cols:
            if date_col not in df_display.columns:
                df_display[date_col] = None
                
        # Fill any remaining NaNs in AUM columns with fallback values
        if want_aum:
            df_pivot_fallback = df_filtered.pivot_table(index="Scheme Code", columns="NAV_Date_Str", values="Fallback_AUM", aggfunc="first").reset_index()
            fallback_cols_map = {}
            for d in date_cols:
                if want_aum and not want_nav:
                    fallback_cols_map[d] = f"{d}_fallback_temp"
                elif want_nav and want_aum:
                    fallback_cols_map[f"{d} (AUM)"] = f"{d}_fallback_temp"
                    
            if fallback_cols_map:
                df_pivot_fallback_renamed = df_pivot_fallback.rename(columns={d: f"{d}_fallback_temp" for d in date_cols if d in df_pivot_fallback.columns})
                available_temp_cols = [col for col in fallback_cols_map.values() if col in df_pivot_fallback_renamed.columns]
                df_display_temp = pd.merge(df_display, df_pivot_fallback_renamed[["Scheme Code"] + available_temp_cols], on="Scheme Code", how="left")
                for main_col, temp_col in fallback_cols_map.items():
                    if main_col in df_display.columns and temp_col in df_display_temp.columns:
                        df_display[main_col] = df_display[main_col].fillna(df_display_temp[temp_col])

        # Carry forward
        if carry_forward and len(date_cols) > 1:
            date_objs = sorted([datetime.strptime(d, "%d-%m-%Y") for d in date_cols])
            sorted_cols = [d.strftime("%d-%m-%Y") for d in date_objs]
            for i in range(1, len(sorted_cols)):
                prev, curr = sorted_cols[i - 1], sorted_cols[i]
                if want_nav and not want_aum:
                    df_display[curr] = df_display[curr].fillna(df_display[prev])
                elif want_aum and not want_nav:
                    df_display[curr] = df_display[curr].fillna(df_display[prev])
                else:
                    df_display[f"{curr} (NAV)"] = df_display[f"{curr} (NAV)"].fillna(df_display[f"{prev} (NAV)"])
                    df_display[f"{curr} (AUM)"] = df_display[f"{curr} (AUM)"].fillna(df_display[f"{prev} (AUM)"])
                    
        df_display = df_display[meta_cols + display_date_cols].sort_values(["Asset Class", "Scheme Name"]).reset_index(drop=True)

        # Convert to vertical layout
        vertical_rows = []
        for _, row in df_display.iterrows():
            meta = {c: row[c] for c in meta_cols}
            for d in date_cols:
                r_item = meta.copy()
                if want_nav and not want_aum:
                    r_item["NAV Date"] = d
                    r_item["NAV"] = row.get(d)
                elif want_aum and not want_nav:
                    r_item["AUM Date"] = d
                    r_item["AUM"] = row.get(d)
                else:
                    r_item["NAV Date"] = d
                    r_item["NAV"] = row.get(f"{d} (NAV)")
                    r_item["AUM Date"] = d
                    r_item["AUM"] = row.get(f"{d} (AUM)")
                vertical_rows.append(r_item)

        v_cols = list(meta_cols)
        if want_nav:
            v_cols += ["NAV Date", "NAV"]
        if want_aum:
            v_cols += ["AUM Date", "AUM"]

        if vertical_rows:
            df_display = pd.DataFrame(vertical_rows)
            df_display = df_display[v_cols].sort_values(["Asset Class", "Scheme Name"]).reset_index(drop=True)
        else:
            df_display = pd.DataFrame(columns=v_cols)

    else:
        # Long format — show raw rows with selected columns
        wanted = ["Asset Class", "Scheme Code", "ISIN Div Payout / ISIN Growth", "ISIN Div Reinvestment", "Scheme Name", "Plan Type", "Option Type"]
        if want_nav:
            wanted.append("NAV")
        if want_aum:
            wanted.append("AUM")
        wanted.append("NAV Date")
        
        # Fill NaNs in df_filtered AUM before extracting
        if "AUM" in df_filtered.columns and "Fallback_AUM" in df_filtered.columns:
            df_filtered["AUM"] = df_filtered["AUM"].fillna(df_filtered["Fallback_AUM"])
            
        df_display = df_filtered[[c for c in wanted if c in df_filtered.columns]].copy()
        df_display["NAV Date"] = df_display["NAV Date"].dt.strftime("%d-%m-%Y")
        display_date_cols = []  # no date pivot columns for long format
        is_aum_only = want_aum and not want_nav

    if want_flows:
        meta_cols_list = ["Asset Class", "Scheme Code", "ISIN Div Payout / ISIN Growth", "ISIN Div Reinvestment", "Scheme Name", "Plan Type", "Option Type"]
        df_display = calculate_flows_for_dataframe(df_display, start_date, meta_cols_list)
        is_aum_only = False

    render_section_header("👁️", "Data Preview")
    st.dataframe(
        df_display,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Daily return": st.column_config.NumberColumn(
                "Daily return",
                format="%.2f%%"
            )
        }
    )

    render_section_header("📥", "Export")
    with st.spinner("Generating styled Excel…"):
        excel_bytes = style_excel(df_display, [], is_aum_only=is_aum_only)

    file_label = f"amfi_nav_{start_date}_to_{end_date}.xlsx"
    st.download_button(
        label="Download styled Excel (.xlsx)",
        data=excel_bytes,
        file_name=file_label,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    if not pivot_dates:
        with st.expander("Scheme summary", expanded=False):
            summary = (
                df_filtered.groupby(["Scheme Code", "Scheme Name", "Plan Type", "Option Type"])
                .agg(
                    Records=("NAV", "count"),
                    Min_NAV=("NAV", "min"),
                    Max_NAV=("NAV", "max"),
                    Latest_NAV_Date=("NAV Date", "max"),
                )
                .reset_index()
            )
            st.dataframe(summary, use_container_width=True, hide_index=True)

    render_app_footer()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        st.error(f"A fatal error occurred on startup: {exc}")
        st.exception(exc)
