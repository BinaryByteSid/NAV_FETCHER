from __future__ import annotations

import os
import re
from datetime import datetime, timedelta
from io import BytesIO, StringIO
from typing import List

import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ─── Robust date parsing (locale-independent) ────────────────────────────────

_MONTH_ABBR = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04",
    "may": "05", "jun": "06", "jul": "07", "aug": "08",
    "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}


def _parse_amfi_date_str(date_str: str) -> str | None:
    """Convert a single AMFI date string like '01-Jun-2026' → '2026-06-01' (ISO)."""
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
    result = pd.to_datetime(series, format="%d-%b-%Y", errors="coerce")
    if not result.isna().all():
        mask = result.isna() & series.notna()
        if mask.any():
            iso_strs = series[mask].apply(_parse_amfi_date_str)
            result[mask] = pd.to_datetime(iso_strs, errors="coerce")
        return result
    result = pd.to_datetime(series, format="%d-%m-%Y", errors="coerce")
    if not result.isna().all():
        return result
    iso_strs = series.apply(_parse_amfi_date_str)
    result = pd.to_datetime(iso_strs, errors="coerce")
    if not result.isna().all():
        return result
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
                    # Log validation error
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


def find_matching_perf_row(nav_name: str, perf_rows: list) -> dict | None:
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


def populate_actual_aum(df: pd.DataFrame, df_port: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()
        
    # 1. Calculate the fallback AUM (NAV-scaled monthly AUM) for every row.
    # This ensures we always have a meaningful, date-varying AUM estimate.
    raw_rows = []
    for idx, r_dict in df.iterrows():
        r_copy = r_dict.copy()
        m_aum = calculate_aum_for_row(r_copy.to_dict(), df_port)
        r_copy["Monthly_AUM"] = m_aum
        raw_rows.append(r_copy)
    df_res = pd.DataFrame(raw_rows)
    
    # Scale the monthly AUM by each day's NAV relative to the scheme mean NAV
    # so that AUM varies day-to-day proportional to NAV movement.
    mean_navs = df_res.groupby("Scheme Code")["NAV"].transform("mean")
    mean_navs = mean_navs.fillna(1.0).replace(0.0, 1.0)
    df_res["Fallback_AUM"] = (df_res["Monthly_AUM"] * (df_res["NAV"] / mean_navs)).round(4)
    
    # Initialize AUM with the fallback so every row already has a value.
    # Rows where the live API succeeds will get overwritten below.
    # Initialize AUM as NaN; we'll fill only where real API values are available.
    df_res["AUM"] = pd.NA
    
    # 2. Try to fetch real AUM from the AMFI performance API for each unique (date, asset-class) pair.
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
    
    # Only fetch for unique (asset_class, date) combinations to minimise API calls.
    unique_groups = df_res[["Asset Class", "Date_Str_Temp"]].drop_duplicates()
    
    perf_lookup = {}
    import time
    
    for i, (_, grp) in enumerate(unique_groups.iterrows()):
        asset_class = grp["Asset Class"]
        date_str = grp["Date_Str_Temp"]
        if not asset_class or not date_str:
            continue
            
        m_id, c_id, s_id = map_section_to_ids(asset_class)
        
        if i > 0:
            time.sleep(0.5)  # rate-limit
            
        perf_rows = fetch_performance_data_from_api(date_str, m_id, c_id, s_id)
        if perf_rows:
            perf_lookup[(date_str, asset_class)] = perf_rows
            
    # Overwrite AUM with real API values where a match is found.
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



def generate_historical_excel(df_final: pd.DataFrame, target_dates: List[str], is_aum_only: bool = False) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="NAV Data")
        
        workbook = writer.book
        worksheet = writer.sheets["NAV Data"]
        
        font_family = "Segoe UI"
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")  # Navy Blue
        header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
        
        data_font = Font(name=font_family, size=10)
        
        # Zebra striping
        even_row_fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
        odd_row_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        thin_border = Border(
            left=Side(style="thin", color="D3D3D3"),
            right=Side(style="thin", color="D3D3D3"),
            top=Side(style="thin", color="D3D3D3"),
            bottom=Side(style="thin", color="D3D3D3")
        )
        
        alignments = {
            "Asset Class": Alignment(horizontal="left", vertical="center"),
            "Scheme Code": Alignment(horizontal="center", vertical="center"),
            "ISIN Div Payout / ISIN Growth": Alignment(horizontal="center", vertical="center"),
            "ISIN Div Reinvestment": Alignment(horizontal="center", vertical="center"),
            "Scheme Name": Alignment(horizontal="left", vertical="center"),
            "Plan Type": Alignment(horizontal="center", vertical="center"),
            "Option Type": Alignment(horizontal="center", vertical="center"),
            "NAV Date": Alignment(horizontal="center", vertical="center"),
            "AUM Date": Alignment(horizontal="center", vertical="center"),
            "NAV": Alignment(horizontal="right", vertical="center"),
            "NAVs": Alignment(horizontal="right", vertical="center"),
            "AUM": Alignment(horizontal="right", vertical="center"),
            "Closing AUM as on previous day": Alignment(horizontal="right", vertical="center"),
            "Actual AUM as on current date": Alignment(horizontal="right", vertical="center"),
            "Daily return": Alignment(horizontal="right", vertical="center"),
            "Derived AUM as on curent day": Alignment(horizontal="right", vertical="center"),
            "Net flows on current day": Alignment(horizontal="right", vertical="center"),
        }
            
        worksheet.row_dimensions[1].height = 28
        for col_idx, col_name in enumerate(df_final.columns, 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center" if col_name not in ["Scheme Name", "Asset Class"] else "left", vertical="center")
            cell.border = thin_border
            
        for row_idx in range(2, len(df_final) + 2):
            worksheet.row_dimensions[row_idx].height = 20
            fill = even_row_fill if row_idx % 2 == 0 else odd_row_fill
            
            for col_idx, col_name in enumerate(df_final.columns, 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.fill = fill
                cell.font = data_font
                cell.border = thin_border
                
                val = cell.value
                if "NAV" in col_name:
                    if val is not None and val != "":
                        cell.number_format = "0.00" if "NAVs" in col_name else "0.0000"
                    else:
                        cell.value = "-"
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                elif "Return" in col_name or "return" in col_name:
                    if val is not None and val != "":
                        cell.number_format = '0.00"%"'
                    else:
                        cell.value = "-"
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                elif "AUM" in col_name or "flows" in col_name or "AUM" in col_name:
                    if val is not None and val != "":
                        cell.number_format = "#,##0.00"
                    else:
                        cell.value = "-"
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                elif "Info Ratio" in col_name:
                    if val is not None and val != "":
                        cell.number_format = "0.0000"
                    else:
                        cell.value = "-"
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                elif "Date" in col_name:
                    if val is None or val == "":
                        cell.value = "-"
                
                align = alignments.get(col_name, Alignment(horizontal="left", vertical="center"))
                cell.alignment = align

        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if cell.row == 1:
                    val_str = val_str + "   "
                if len(val_str) > max_len:
                    max_len = len(val_str)
            worksheet.column_dimensions[col_letter].width = max(min(max_len + 2, 55), 12)
            
    return buffer.getvalue()

from amfi_nav import (
    AMFINavError,
    benchmark_delta,
    comparison_table,
    export_to_excel,
    fetch_nav_data,
    filter_family_rows,
    load_historical_snapshots,
    search_fund,
    sip_future_value,
    summarize_families,
    classify_plan_type,
    classify_option_type,
)
from nav_fetcher import parse_bucket_input, run_live_portfolio, style_portfolio_excel
from ui_theme import (
    finance_panel,
    inject_custom_css,
    render_app_footer,
    render_hero,
    render_info_card,
    render_section_header,
    render_ticker_strip,
    render_top_bar,
)


st.set_page_config(
    page_title="NAV Terminal · AMFI India",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def frame_to_excel_bytes(frame: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    export_to_excel(frame, buffer)
    return buffer.getvalue()


def frame_to_csv_bytes(frame: pd.DataFrame) -> bytes:
    buffer = StringIO()
    export_frame = frame.copy()
    if "NAV Date" in export_frame.columns:
        export_frame["NAV Date"] = pd.to_datetime(export_frame["NAV Date"], errors="coerce").dt.strftime("%d-%m-%Y")
    export_frame.to_csv(buffer, index=False)
    return buffer.getvalue().encode("utf-8")


def build_export_frame(frame: pd.DataFrame) -> pd.DataFrame:
    export_frame = frame.copy()
    columns = [
        "Scheme Name",
        "Scheme Code",
        "Plan Type",
        "Option Type",
        "NAV",
        "NAV Date",
        "AMC Name",
    ]
    available = [column for column in columns if column in export_frame.columns]
    export_frame = export_frame[available]
    export_frame = export_frame.rename(columns={"Option Type": "Plan Option"})
    return export_frame


def normalize_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9_-]+", "_", value)
    return value.strip("_")[:80]


@st.cache_data(ttl=3600, show_spinner=False)
def load_data(force_refresh: bool = False) -> pd.DataFrame:
    return fetch_nav_data(force_refresh=force_refresh)


def format_timestamp(frame: pd.DataFrame) -> str:
    if frame.empty or "NAV Date" not in frame.columns:
        return "Unavailable"
    latest = pd.to_datetime(frame["NAV Date"], errors="coerce").dropna().max()
    if pd.isna(latest):
        return "Unavailable"
    return pd.Timestamp(latest).strftime("%d-%m-%Y")


def render_result_card(selected_summary: pd.Series, selected_rows: pd.DataFrame) -> None:
    render_section_header("📈", "Fund Overview", "Key metrics for the selected scheme family")
    latest_nav = selected_summary.get("Latest NAV")
    nav_display = f"{latest_nav:.4f}" if pd.notna(latest_nav) else "N/A"
    latest_date = pd.to_datetime(selected_summary.get("Latest NAV Date"), errors="coerce")
    date_display = latest_date.strftime("%d-%m-%Y") if pd.notna(latest_date) else "N/A"
    st.markdown(
        f"""
        <div class="metrics-grid">
            <div class="metric-card">
                <span class="metric-label">Fund</span>
                <span class="metric-value">{selected_summary.get("Family Name", "-")}</span>
            </div>
            <div class="metric-card">
                <span class="metric-label">AMC</span>
                <span class="metric-value">{selected_summary.get("AMC Name", "-")}</span>
            </div>
            <div class="metric-card">
                <span class="metric-label">Latest NAV</span>
                <span class="metric-value accent">{nav_display}</span>
            </div>
            <div class="metric-card">
                <span class="metric-label">NAV Date</span>
                <span class="metric-value">{date_display}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_section_header("⚖️", "Regular vs Direct Comparison")
    comparison = comparison_table(selected_rows, selected_summary["Family Key"])
    if comparison.empty:
        st.info("No comparison rows available for the selected fund.")
    else:
        st.dataframe(comparison, use_container_width=True, hide_index=True)


def render_history(selected_summary: pd.Series) -> None:
    render_section_header("📉", "Historical NAV Cache", "Trends from locally cached snapshots")
    history = load_historical_snapshots(selected_summary["Family Key"])
    if history.empty:
        st.info("No cached historical snapshots are available yet. The app will build them after a successful refresh.")
        return

    timeline = history[["NAV Date", "Scheme Name", "Plan Type", "Option Type", "NAV"]].copy()
    timeline["NAV Date"] = pd.to_datetime(timeline["NAV Date"], errors="coerce")
    timeline = timeline.dropna(subset=["NAV Date"])
    st.line_chart(timeline.pivot_table(index="NAV Date", values="NAV", aggfunc="mean"))
    st.dataframe(timeline.sort_values("NAV Date", ascending=False), use_container_width=True, hide_index=True)

    render_section_header("🎯", "Benchmark Comparison")
    benchmark_return = st.number_input("Assumed benchmark annual return %", min_value=0.0, value=10.0, step=0.5)
    daily_series = timeline.groupby("NAV Date", as_index=False)["NAV"].mean().sort_values("NAV Date")
    delta = benchmark_delta(daily_series["NAV"], benchmark_return)
    metric_cols = st.columns(3)
    metric_cols[0].metric("Observed change %", "N/A" if delta["observed_change_pct"] is None else f"{delta['observed_change_pct']:.2f}")
    metric_cols[1].metric("Benchmark change %", "N/A" if delta["benchmark_change_pct"] is None else f"{delta['benchmark_change_pct']:.2f}")
    metric_cols[2].metric("Delta %", "N/A" if delta["delta_pct"] is None else f"{delta['delta_pct']:.2f}")


def render_sip_calculator(default_nav: float | None) -> None:
    render_section_header("🧮", "SIP Calculator", "Project future corpus from monthly investments")
    col1, col2, col3 = st.columns(3)
    with col1:
        monthly = st.number_input("Monthly SIP amount", min_value=0.0, value=5000.0, step=500.0)
    with col2:
        annual_return = st.number_input("Expected annual return %", min_value=0.0, value=12.0, step=0.5)
    with col3:
        years = st.number_input("Investment horizon (years)", min_value=1, value=5, step=1)

    future_value = sip_future_value(monthly, annual_return, int(years))
    invested = monthly * 12 * years
    st.metric("Projected corpus", f"₹{future_value:,.2f}")
    st.caption(f"Total invested: ₹{invested:,.2f}")
    if default_nav and default_nav > 0:
        st.caption(f"Reference NAV used for display only: {default_nav:.4f}")


def main() -> None:
    inject_custom_css()
    render_top_bar()

    try:
        nav_data = load_data(force_refresh=False)
    except Exception as exc:
        st.error(f"Fatal error loading data: {exc}")
        st.stop()

    latest_update = format_timestamp(nav_data)
    scheme_count = len(nav_data) if not nav_data.empty else 0
    render_ticker_strip()
    render_hero(
        title="Mutual Fund NAV Terminal",
        subtitle=(
            "Institutional-grade search, comparison, and export for Indian mutual funds — "
            "live AMFI feeds, fuzzy matching, Regular vs Direct analysis, and styled Excel reports."
        ),
        chips=[
            {"label": "Latest NAV Date", "value": latest_update, "tone": "positive"},
            {"label": "Schemes Indexed", "value": f"{scheme_count:,}", "tone": "gold"},
            {"label": "Data Source", "value": "AMFI India", "tone": ""},
            {"label": "Market", "value": "Mutual Funds", "tone": ""},
        ],
    )

    with finance_panel("Search & Export"):
        render_section_header("🔍", "Fund Discovery", "Search single funds, run batch lookups, or generate historical ISIN reports")
        search_mode = st.radio(
            "Search mode",
            ["Single Fund", "Batch Search", "Historical ISIN Export", "Fund Performance", "Portfolio Bucket Tracker"],
            horizontal=True,
            index=0,
            label_visibility="collapsed",
        )

        selected_rows = pd.DataFrame()
        batch_export_rows = pd.DataFrame()
        selected_summary = None

        if search_mode == "Single Fund":
            query = st.text_input(
                "Search funds",
                placeholder="e.g. HDFC Flexi Cap or 120439",
                help="Fuzzy search by fund name or exact Scheme Code",
            )
            c_refresh, c_latest = st.columns([1, 1])
            with c_refresh:
                refresh_requested = st.button("Refresh from AMFI", help="Fetch the latest official feed once and update the cache.")
            with c_latest:
                st.metric("Latest update", latest_update)

            plan_filters = st.multiselect("Plan filter", ["Regular", "Direct"], default=["Regular", "Direct"])
            option_filters = st.multiselect("Option filter", ["Growth", "IDCW/Dividend", "Bonus", "Other"], default=["Growth", "IDCW/Dividend"])

            if refresh_requested:
                try:
                    nav_data = load_data(force_refresh=True)
                    summary = summarize_families(nav_data)
                    latest_update = format_timestamp(nav_data)
                    st.success("Refreshed AMFI data.")
                except Exception as exc:
                    st.error(f"Failed to refresh data: {exc}")
                    st.stop()

            suggestions: List[str] = []
            if query.strip():
                matches = search_fund(nav_data, query, plan_filter=plan_filters, option_filter=option_filters, limit=10)
                suggestions = matches["Family Name"].tolist() if not matches.empty else []

            selected_name = None
            if suggestions:
                selected_name = st.selectbox("Matching results", suggestions, index=0)
            elif query.strip():
                st.warning("No fuzzy matches found. Try a broader fund name.")

            if selected_name:
                matched = search_fund(nav_data, selected_name, plan_filter=plan_filters, option_filter=option_filters, limit=1)
                if not matched.empty:
                    selected_summary = matched.iloc[0]
                    selected_rows = filter_family_rows(nav_data, selected_summary["Family Key"], plan_filter=plan_filters, option_filter=option_filters)

        elif search_mode == "Batch Search":
            c_refresh, c_latest = st.columns([1, 1])
            with c_refresh:
                refresh_requested = st.button("Refresh from AMFI", help="Fetch the latest official feed once and update the cache.")
            with c_latest:
                st.metric("Latest update", latest_update)

            plan_filters = st.multiselect("Plan filter", ["Regular", "Direct"], default=["Regular", "Direct"])
            option_filters = st.multiselect("Option filter", ["Growth", "IDCW/Dividend", "Bonus", "Other"], default=["Growth", "IDCW/Dividend"])

            if refresh_requested:
                try:
                    nav_data = load_data(force_refresh=True)
                    summary = summarize_families(nav_data)
                    latest_update = format_timestamp(nav_data)
                    st.success("Refreshed AMFI data.")
                except Exception as exc:
                    st.error(f"Failed to refresh data: {exc}")
                    st.stop()

            render_info_card(
                "<strong>Batch mode:</strong> Enter multiple fund names or scheme codes — "
                "one per line or comma-separated — to search and export results in one go."
            )
            batch_queries = st.text_area(
                "Batch search",
                placeholder="HDFC Equity\nAxis Bluechip\n120439",
                height=120,
            )
            if st.button("Run batch search", type="primary") and batch_queries.strip():
                query_items = [item.strip() for item in re.split(r"[,\n]+", batch_queries) if item.strip()]
                batch_matches = []
                batch_rows = []
                for query in query_items:
                    found = search_fund(nav_data, query, plan_filter=plan_filters, option_filter=option_filters, limit=5)
                    if not found.empty:
                        found = found.copy()
                        found.insert(0, "Search Query", query)
                        batch_matches.append(found)
                        batch_rows.extend(
                            filter_family_rows(nav_data, family_key, plan_filter=plan_filters, option_filter=option_filters)
                            for family_key in found["Family Key"].dropna().tolist()
                        )

                if batch_matches:
                    selected_rows = pd.concat(batch_matches, ignore_index=True)
                    flattened_rows = [row for row in batch_rows if not row.empty]
                    if flattened_rows:
                        batch_export_rows = pd.concat(flattened_rows, ignore_index=True).drop_duplicates(
                            subset=["Family Key", "Scheme Code", "NAV Date", "Plan Type", "Option Type"]
                        )
                    st.dataframe(selected_rows, use_container_width=True, hide_index=True)
                else:
                    st.warning("No matching funds found for the batch search.")

        elif search_mode == "Fund Performance":
            # UI for Fund Performance dashboard matching the client's design
            maturity_type = st.selectbox("Maturity Type", ["Open Ended", "Close Ended", "Interval"], index=0)
            category = st.selectbox("Category", ["Equity", "Debt", "Hybrid", "Solution Oriented", "Other"], index=0)
            maturity_id_map = {"Open Ended": 1, "Close Ended": 2, "Interval": 2}
            cat_id_map = {"Equity": 1, "Debt": 2, "Hybrid": 3, "Solution Oriented": 4, "Other": 5}
            maturity_id = maturity_id_map[maturity_type]
            cat_id = cat_id_map[category]

            # Fetch subcategories using the correct keys 'name' and 'id'
            subcategories = []
            with st.spinner("Fetching subcategories..."):
                try:
                    sub_resp = requests.post(
                        "https://www.amfiindia.com/gateway/pollingsebi/api/amfi/getsubcategory",
                        json={"category": cat_id},
                        headers={"User-Agent": "Mozilla/5.0", "Content-Type": "application/json"},
                        timeout=20,
                    )
                    if sub_resp.status_code == 200:
                        sub_data = sub_resp.json()
                        subcategories = [(item.get("name"), item.get("id")) for item in sub_data.get("data", []) if item.get("name")]
                except Exception as e:
                    st.error(f"Failed to fetch subcategories: {e}")

            subcategory_name = st.selectbox("Subcategory", [name for name, _ in subcategories] or ["All"], index=0)
            sub_id = dict(subcategories).get(subcategory_name, 0)

            # Return Period selector and date picker
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                horizon = st.selectbox("Return Horizon", ["1-Year", "3-Year", "5-Year", "10-Year", "Since Inception"], index=2)
            with col_d2:
                report_date = st.date_input("Report Date", value=datetime.today().date())

            if st.button("Go", type="primary", use_container_width=True):
                with st.spinner("Fetching performance data..."):
                    date_str = report_date.strftime("%d-%b-%Y")
                    perf_rows = fetch_performance_data_from_api(date_str, maturity_id, cat_id, sub_id)
                    if not perf_rows:
                        # Auto-fallback: search backwards for a date that has data
                        found_date = None
                        with st.spinner("Checking latest available date on AMFI..."):
                            for i in range(1, 15):
                                check_date = report_date - timedelta(days=i)
                                check_str = check_date.strftime("%d-%b-%Y")
                                check_rows = fetch_performance_data_from_api(check_str, maturity_id, cat_id, sub_id)
                                if check_rows:
                                    found_date = check_date
                                    perf_rows = check_rows
                                    break
                        if found_date:
                            st.warning(f"⚠️ No performance data is available for {report_date.strftime('%d-%b-%Y')}. "
                                       f"AMFI updates performance data weekly. Showing the latest available report from **{found_date.strftime('%A, %d-%b-%Y')}**.")
                            report_date = found_date
                            date_str = found_date.strftime("%d-%b-%Y")
                        else:
                            st.warning(f"⚠️ No performance data returned for the selected criteria on {report_date.strftime('%d-%b-%Y')} or the last 14 days.")
                    
                    if perf_rows:
                        df_perf = pd.DataFrame(perf_rows)
                        
                        horizon_map = {
                            "1-Year": ("return1YearRegular", "return1YearDirect", "return1YearBenchmark", "ir1YrRegular", "ir1YrDirect", "1-Year"),
                            "3-Year": ("return3YearRegular", "return3YearDirect", "return3YearBenchmark", "ir3YrRegular", "ir3YrDirect", "3-Year"),
                            "5-Year": ("return5YearRegular", "return5YearDirect", "return5YearBenchmark", "ir5YrRegular", "ir5YrDirect", "5-Year"),
                            "10-Year": ("return10YearRegular", "return10YearDirect", "return10YearBenchmark", "ir10YrRegular", "ir10YrDirect", "10-Year"),
                            "Since Inception": ("returnSinceLaunchRegular", "returnSinceLaunchDirect", "returnSinceLaunchBenchmarkRegular", None, None, "Since Inception")
                        }
                        reg_ret, dir_ret, bench_ret, reg_ir, dir_ir, label = horizon_map[horizon]
                        
                        # Build formatted DataFrame for display
                        rename_dict = {
                            "schemeName": "Scheme Name",
                            "benchmark": "Benchmark Name",
                            "riskometerScheme": "Riskometer (Scheme)",
                            "riskometerBenchmark": "Riskometer (Benchmark)",
                            "navRegular": "Latest NAV (Regular)",
                            "navDirect": "Latest NAV (Direct)",
                            reg_ret: f"{label} Return (Regular)",
                            dir_ret: f"{label} Return (Direct)",
                            bench_ret: f"{label} Return (Benchmark)",
                        }
                        if reg_ir and reg_ir in df_perf.columns:
                            rename_dict[reg_ir] = f"{label} Info Ratio (Regular)"
                        if dir_ir and dir_ir in df_perf.columns:
                            rename_dict[dir_ir] = f"{label} Info Ratio (Direct)"
                        rename_dict["dailyAUM"] = "Daily AUM (Cr.)"
                        
                        existing_renames = {k: v for k, v in rename_dict.items() if k in df_perf.columns}
                        df_display = df_perf[list(existing_renames.keys())].rename(columns=existing_renames)
                        
                        # Color coding riskometer scheme/benchmark
                        def highlight_riskometer(val):
                            if val == "Very High":
                                return "background-color: #fca5a5; color: #991b1b; font-weight: bold; text-align: center;"
                            elif val == "High":
                                return "background-color: #fed7aa; color: #9a3412; font-weight: bold; text-align: center;"
                            elif val == "Moderate":
                                return "background-color: #fef08a; color: #854d0e; font-weight: bold; text-align: center;"
                            elif val == "Low to Moderate":
                                return "background-color: #d9f99d; color: #3f6212; font-weight: bold; text-align: center;"
                            elif val == "Low":
                                return "background-color: #bbf7d0; color: #166534; font-weight: bold; text-align: center;"
                            return ""

                        styled_df = df_display.style
                        subset_cols = [col for col in ["Riskometer (Scheme)", "Riskometer (Benchmark)"] if col in df_display.columns]
                        if hasattr(styled_df, "map"):
                            styled_df = styled_df.map(highlight_riskometer, subset=subset_cols)
                        else:
                            styled_df = styled_df.applymap(highlight_riskometer, subset=subset_cols)
                        
                        # Display the beautiful interactive styled table
                        st.dataframe(
                            styled_df,
                            use_container_width=True,
                            column_config={
                                "Latest NAV (Regular)": st.column_config.NumberColumn(format="₹%.4f"),
                                "Latest NAV (Direct)": st.column_config.NumberColumn(format="₹%.4f"),
                                f"{label} Return (Regular)": st.column_config.NumberColumn(format="%.2f%%"),
                                f"{label} Return (Direct)": st.column_config.NumberColumn(format="%.2f%%"),
                                f"{label} Return (Benchmark)": st.column_config.NumberColumn(format="%.2f%%"),
                                f"{label} Info Ratio (Regular)": st.column_config.NumberColumn(format="%.4f"),
                                f"{label} Info Ratio (Direct)": st.column_config.NumberColumn(format="%.4f"),
                                "Daily AUM (Cr.)": st.column_config.NumberColumn(format="₹%,.2f Cr"),
                            }
                        )
                        
                        # Styled excel report
                        excel_bytes = generate_historical_excel(df_display, [], is_aum_only=False)
                        st.download_button(
                            label="📥 Download Styled Excel (.xlsx)",
                            data=excel_bytes,
                            file_name=f"fund_performance_{subcategory_name}_{date_str}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
        elif search_mode == "Portfolio Bucket Tracker":
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
                
            if "portfolio_results" not in st.session_state:
                st.session_state["portfolio_results"] = None
            if "last_active_bucket" not in st.session_state:
                st.session_state["last_active_bucket"] = st.session_state["active_bucket_name"]
            if "last_processed_file" not in st.session_state:
                st.session_state["last_processed_file"] = None
                
            # If active bucket changed, reset portfolio results to avoid stale data display
            if st.session_state["last_active_bucket"] != st.session_state["active_bucket_name"]:
                st.session_state["portfolio_results"] = None
                st.session_state["last_active_bucket"] = st.session_state["active_bucket_name"]

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
                new_bucket_name_raw = st.text_input("New Bucket Name", placeholder="e.g. My Conservative Portfolio")
                new_bucket_name = new_bucket_name_raw.strip()
                c_add, c_del = st.columns(2)
                with c_add:
                    if st.button("Create New Bucket", use_container_width=True):
                        if new_bucket_name and new_bucket_name not in st.session_state["portfolio_buckets"]:
                            curr_df = st.session_state["portfolio_buckets"][st.session_state["active_bucket_name"]].copy()
                            st.session_state["portfolio_buckets"][new_bucket_name] = curr_df
                            st.session_state["active_bucket_name"] = new_bucket_name
                            st.session_state["portfolio_results"] = None
                            st.rerun()
                with c_del:
                    if st.button("Delete Active Bucket", use_container_width=True, type="secondary"):
                        if len(st.session_state["portfolio_buckets"]) > 1:
                            del st.session_state["portfolio_buckets"][st.session_state["active_bucket_name"]]
                            st.session_state["active_bucket_name"] = list(st.session_state["portfolio_buckets"].keys())[0]
                            st.session_state["portfolio_results"] = None
                            st.rerun()
                        else:
                            st.warning("Cannot delete the last remaining bucket.")

            st.markdown('<hr class="panel-divider">', unsafe_allow_html=True)
            
            # 2. Investment Setup
            render_section_header("📅", "Investment Setup", "Pick when you started investing — the tracker runs to today automatically")
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                start_date = st.date_input("Investment Start Date", value=(datetime.today() - timedelta(days=90)).date())
            with col_b:
                initial_amount_str = st.text_input("Investment Amount (₹)", value="100000")
                try:
                    initial_amount = float(initial_amount_str.replace(",", "").strip())
                    if initial_amount <= 0:
                        st.error("Investment Amount must be greater than 0.")
                        initial_amount = 100000.0
                except ValueError:
                    st.error("Please enter a valid numeric Investment Amount.")
                    initial_amount = 100000.0
            with col_c:
                skip_sunday = st.checkbox("Skip Sundays", value=True)

            st.markdown('<hr class="panel-divider">', unsafe_allow_html=True)
            
            # 3. Bucket Composition Editor
            render_section_header("📋", "Bucket Composition", f"Add/remove funds and set weights for '{st.session_state['active_bucket_name']}'")
            
            uploaded_file = st.file_uploader("Upload Bucket (Excel/CSV with ISIN and Weights)", type=["xlsx", "xls", "csv"])
            
            current_bucket_df = st.session_state["portfolio_buckets"][st.session_state["active_bucket_name"]]
            if uploaded_file is not None:
                file_key = f"processed_{uploaded_file.name}_{uploaded_file.size}"
                if st.session_state.get("last_processed_file") != file_key:
                    parsed_df = parse_bucket_input(uploaded_file=uploaded_file)
                    if not parsed_df.empty:
                        st.session_state["portfolio_buckets"][st.session_state["active_bucket_name"]] = parsed_df
                        current_bucket_df = parsed_df
                        st.session_state["portfolio_results"] = None
                        st.session_state["last_processed_file"] = file_key
                        st.success("Successfully loaded bucket from file!")
            else:
                st.session_state["last_processed_file"] = None
                    
            current_bucket_df_editor = current_bucket_df.copy()
            if "Weight (%)" in current_bucket_df_editor.columns:
                current_bucket_df_editor["Weight (%)"] = current_bucket_df_editor["Weight (%)"].astype(str)
                
            edited_df = st.data_editor(
                current_bucket_df_editor,
                column_config={
                    "Scheme Name": st.column_config.TextColumn("Scheme Name", width="large", help="Optional name for display"),
                    "ISIN": st.column_config.TextColumn("ISIN", required=True, help="Mutual Fund ISIN (Growth or Reinvestment)"),
                    "Weight (%)": st.column_config.TextColumn("Weight (%)", required=True, help="Percentage weight in portfolio (e.g. 8.9123456789)")
                },
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_{st.session_state['active_bucket_name']}"
            )
            
            if edited_df is not None:
                # Convert manual string weight input back to float, replacing invalid text with 0.0
                if "Weight (%)" in edited_df.columns:
                    edited_df["Weight (%)"] = pd.to_numeric(
                        edited_df["Weight (%)"].astype(str).str.replace("%", "").str.strip(), 
                        errors="coerce"
                    ).fillna(0.0)
                st.session_state["portfolio_buckets"][st.session_state["active_bucket_name"]] = edited_df
                
            weight_sum = edited_df["Weight (%)"].sum() if not edited_df.empty else 0
            weight_sum_str = f"{weight_sum:.10f}".rstrip('0').rstrip('.')
            if weight_sum != 100.0:
                st.info(f"⚖️ Current weights sum to **{weight_sum_str}%**. Weights will be automatically normalized to 100%.")
            else:
                st.success("⚖️ Weights sum to exactly **100%**!")

            st.markdown('<hr class="panel-divider">', unsafe_allow_html=True)
            
            # 4. Track / Refresh buttons
            col_btn1, col_btn2 = st.columns([3, 1])
            with col_btn1:
                track_btn = st.button("📊 Track Portfolio", type="primary", use_container_width=True)
            with col_btn2:
                refresh_btn = st.button("🔄 Refresh", use_container_width=True)
            
            if track_btn or refresh_btn:
                if start_date > datetime.today().date():
                    st.error("Start Date cannot be in the future.")
                else:
                    with st.spinner("Fetching latest NAVs from AMFI..."):
                        res_temp, err_temp = run_live_portfolio(
                            edited_df,
                            start_date,
                            initial_amount,
                            skip_sunday
                        )
                    if err_temp:
                        st.error(err_temp)
                        st.session_state["portfolio_results"] = None
                    else:
                        st.session_state["portfolio_results"] = res_temp
                        
            if st.session_state["portfolio_results"] is not None:
                res = st.session_state["portfolio_results"]
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
                
                # ── Fund-wise holdings ───────────────────────────────────
                render_section_header("💼", "Fund-wise Holdings")
                st.dataframe(
                    res["composition"],
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Weight (%)": st.column_config.NumberColumn("Weight (%)", format="%.10f%%"),
                        "Units": st.column_config.NumberColumn("Units", format="%.10f"),
                        "Return (%)": st.column_config.NumberColumn("Return (%)", format="%.2f%%"),
                    }
                )
                
                # ── Portfolio Value Chart ─────────────────────────────────
                render_section_header("📈", "Portfolio Value Over Time")
                chart_df = res["tracker"].copy()
                chart_df["Date_dt"] = pd.to_datetime(chart_df["Date"], format="%d-%m-%Y")
                chart_df = chart_df.set_index("Date_dt")
                chart_cols = ["Total Portfolio Value (₹)"]
                if "Nifty 50 Valuation (₹)" in chart_df.columns:
                    chart_cols.append("Nifty 50 Valuation (₹)")
                if "Nifty 500 Valuation (₹)" in chart_df.columns:
                    chart_cols.append("Nifty 500 Valuation (₹)")
                st.line_chart(chart_df[chart_cols])
                
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
        elif search_mode == "Historical ISIN Export":
            render_section_header("📋", "Historical ISIN Export", "Pivoted NAV/AUM reports with corporate Excel styling")
            render_info_card(
                "<strong>Historical NAV Extractor:</strong> Specify a date range and target ISINs to generate "
                "a pivoted, corporate-styled Excel sheet. Weekends and holidays can be filled via carry-forward."
            )

            col1, col2 = st.columns(2)
            with col1:
                start_date = st.date_input("Start Date", value=datetime(2000, 1, 1).date())
            with col2:
                end_date = st.date_input("End Date", value=datetime(2026, 5, 29).date())

            default_isins_str = "\n".join([
                "INF209K01AJ8", "INF846K01CH7", "INF846K016E3", "INF194K01524", "INF760K01019", 
                "INF760K01KR2", "INF740K01128", "INF179K01608", "INF179KA1RT1", "INF179K01CR2", 
                "INF179KA1RZ8", "INF109KA1TX4", "INF109K01BZ4", "INF205KA1189", "INF205K011T7", 
                "INF174KA1EK3", "INF174K01DS9", "INF174KA1HS9", "INF247L01478", "INF204K01GE7", 
                "INF204K01562", "INF204K01489", "INF879O01019", "INF966L01457", "INF966L01AW4", 
                "INF966L01234", "INF200K01370", "INF200K01CT2", "INF200K01297"
            ])
        
            isin_input = st.text_area("Target ISINs (one per line or comma-separated)", value=default_isins_str, height=200)
        
            c1, c2 = st.columns(2)
            with c1:
                carry_forward = st.checkbox("Carry forward NAV on holidays/weekends", value=True)
            with c2:
                skip_sundays = st.checkbox("Skip Sundays", value=True)

            col_c1, col_c2, col_c3 = st.columns(3)
            with col_c1:
                want_nav = st.checkbox("Want NAV", value=True)
            with col_c2:
                want_aum = st.checkbox("Want AUM", value=True)
            with col_c3:
                want_flows = st.checkbox("Want Flows", value=False)

            if st.button("Fetch & generate Excel", type="primary", use_container_width=True):
                parsed_isins = [x.strip() for x in re.split(r"[,\n\s]+", isin_input) if x.strip()]
                if not want_nav and not want_aum and not want_flows:
                    st.error("Please select at least one data type (NAV, AUM, or Flows) to export.")
                elif not parsed_isins:
                    st.error("Please enter at least one valid ISIN.")
                elif start_date > end_date:
                    st.error("Start Date cannot be after End Date.")
                else:
                    if want_flows:
                        want_nav = True
                        want_aum = True
                        
                    fetch_start_date = start_date
                    if want_flows:
                        fetch_start_date = start_date - timedelta(days=10)
                        
                    with st.spinner("Connecting to AMFI India and fetching historical data..."):
                        try:
                            frmdt_str = fetch_start_date.strftime("%d-%b-%Y")
                            todt_str = end_date.strftime("%d-%b-%Y")
                            
                            url = f"https://portal.amfiindia.com/DownloadNAVHistoryReport_Po.aspx?frmdt={frmdt_str}&todt={todt_str}"
                            import time
                            max_retries = 3
                            delay = 2.0
                            response = None
                            for attempt in range(max_retries):
                                try:
                                    response = requests.get(url, stream=True, timeout=300)
                                    response.raise_for_status()
                                    break
                                except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                                    if attempt == max_retries - 1:
                                        raise e
                                    time.sleep(delay)
                                    delay *= 2
                        
                            rows = []
                            current_section = "Unknown"
                        
                            df_port = load_portfolio_aum_data()
                            parsed_isins_set = {x.upper() for x in parsed_isins}
                        
                            for line_bytes in response.iter_lines():
                                if not line_bytes:
                                    continue
                                line = line_bytes.decode('utf-8', errors='ignore')
                            
                                # Fast check: AMC and section lines do not contain semicolons
                                if ";" not in line:
                                    line_stripped = line.strip()
                                    if (
                                        line_stripped.startswith("Open Ended")
                                        or line_stripped.startswith("Closed Ended")
                                        or line_stripped.startswith("Interval Fund Schemes")
                                    ):
                                        current_section = line_stripped
                                    continue
                            
                                parts = line.split(";")
                                if len(parts) < 8:
                                    continue
                                
                                isin_growth = parts[2].strip()
                                isin_reinvestment = parts[3].strip()
                            
                                isin_growth_upper = isin_growth.upper() if isin_growth != "-" else ""
                                isin_reinvest_upper = isin_reinvestment.upper() if isin_reinvestment != "-" else ""
                            
                                g_match = isin_growth_upper and isin_growth_upper in parsed_isins_set
                                r_match = isin_reinvest_upper and isin_reinvest_upper in parsed_isins_set
                            
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
                        
                            if not rows:
                                st.warning("No NAV records found matching the specified ISINs and date range.")
                            else:
                                df_raw = pd.DataFrame(rows)
                            
                                df_raw["Plan Type"] = df_raw["Scheme Name"].apply(classify_plan_type)
                                df_raw["Option Type"] = df_raw["Scheme Name"].apply(classify_option_type)
                            
                                # Infer proper Asset Class for rows where the Asset Class is a fund family name
                                known_sections = ["Open Ended", "Closed Ended", "Interval Fund Schemes"]
                                def infer_section_from_name(name: str) -> str:
                                    lowered = name.lower()
                                    if "flexi" in lowered:
                                        return "Open Ended Schemes (Equity Scheme - Flexi Cap Fund)"
                                    if "mid" in lowered:
                                        return "Open Ended Schemes (Equity Scheme - Mid Cap Fund)"
                                    if "small" in lowered:
                                        return "Open Ended Schemes (Equity Scheme - Small Cap Fund)"
                                    if "large" in lowered:
                                        return "Open Ended Schemes (Equity Scheme - Large Cap Fund)"
                                    if "focused" in lowered or "focus" in lowered:
                                        return "Open Ended Schemes (Equity Scheme - Focused Fund)"
                                    return "Open Ended Schemes (Equity Scheme - Large Cap Fund)"
                                df_raw["Asset Class"] = df_raw.apply(
                                    lambda row: row["Asset Class"] if any(row["Asset Class"].startswith(s) for s in known_sections) else infer_section_from_name(row["Scheme Name"]),
                                    axis=1)
                                df_raw = populate_actual_aum(df_raw, df_port)
                                if not df_raw.empty:
                                    # Parse dates robustly (locale-independent)
                                    parsed_dates = parse_amfi_date_series(df_raw["Date"])
                                    df_raw["Date_parsed"] = parsed_dates
                                    df_raw["Date"] = parsed_dates.dt.strftime("%d-%m-%Y")
                                    # Sort by scheme + parsed date so shift() gives the correct previous trading day
                                    df_raw = df_raw.sort_values(["Scheme Code", "Date_parsed"]).reset_index(drop=True)
                                    # Compute daily return %: only where NAV actually changed between real trading days
                                    df_raw["Prev NAV"] = df_raw.groupby("Scheme Code")["NAV"].shift(1)
                                    df_raw["Daily Return %"] = (
                                        (df_raw["NAV"] - df_raw["Prev NAV"]) / df_raw["Prev NAV"] * 100
                                    ).where(df_raw["Prev NAV"].notna() & (df_raw["Prev NAV"] != 0))
                                    # Flows = Daily Return % * previous-day AUM / 100
                                    df_raw["Prev AUM"] = df_raw.groupby("Scheme Code")["AUM"].shift(1)
                                    df_raw["Flows"] = (
                                        df_raw["Daily Return %"] / 100 * df_raw["Prev AUM"]
                                    ).where(df_raw["Daily Return %"].notna() & df_raw["Prev AUM"].notna())
                                    df_raw.drop(columns=["Prev NAV", "Prev AUM", "Date_parsed"], inplace=True)
                            
                                all_dates = pd.date_range(start=fetch_start_date, end=end_date)
                            
                                target_dates = []
                                for dt in all_dates:
                                    if skip_sundays and dt.weekday() == 6:
                                        continue
                                    target_dates.append(dt.strftime("%d-%m-%Y"))
                                
                                fund_metadata = df_raw[[
                                    "Asset Class", 
                                    "Scheme Code", 
                                    "ISIN Div Payout / ISIN Growth", 
                                    "ISIN Div Reinvestment", 
                                    "Scheme Name", 
                                    "Plan Type", 
                                    "Option Type"
                                ]].drop_duplicates(subset=["Scheme Code"])
                            
                                if want_nav and not want_aum:
                                    df_pivot = df_raw.pivot(index="Scheme Code", columns="Date", values="NAV").reset_index()
                                    display_date_cols = target_dates
                                    is_aum_only = False
                                elif want_aum and not want_nav:
                                    df_pivot = df_raw.pivot(index="Scheme Code", columns="Date", values="AUM").reset_index()
                                    display_date_cols = target_dates
                                    is_aum_only = True
                                else:
                                    # Pivot NAV, AUM, and Flows
                                    df_pivot_nav = df_raw.pivot(index="Scheme Code", columns="Date", values="NAV").reset_index()
                                    df_pivot_aum = df_raw.pivot(index="Scheme Code", columns="Date", values="AUM").reset_index()
                                    df_pivot_flow = df_raw.pivot(index="Scheme Code", columns="Date", values="Flows").reset_index()
                                    # Rename columns with appropriate suffixes
                                    nav_cols_map = {d: f"{d} (NAV)" for d in target_dates if d in df_pivot_nav.columns}
                                    aum_cols_map = {d: f"{d} (AUM)" for d in target_dates if d in df_pivot_aum.columns}
                                    flow_cols_map = {d: f"{d} (Flows)" for d in target_dates if d in df_pivot_flow.columns}
                                    df_pivot_nav = df_pivot_nav.rename(columns=nav_cols_map)
                                    df_pivot_aum = df_pivot_aum.rename(columns=aum_cols_map)
                                    df_pivot_flow = df_pivot_flow.rename(columns=flow_cols_map)
                                    # Merge all three dataframes
                                    df_pivot = pd.merge(df_pivot_nav, df_pivot_aum, on="Scheme Code", how="left")
                                    df_pivot = pd.merge(df_pivot, df_pivot_flow, on="Scheme Code", how="left")
                                    # Build interleaved display columns including Flows
                                    interleaved_dates = []
                                    for d in target_dates:
                                        interleaved_dates.append(f"{d} (NAV)")
                                        interleaved_dates.append(f"{d} (AUM)")
                                        interleaved_dates.append(f"{d} (Flows)")
                                    display_date_cols = interleaved_dates
                                    is_aum_only = False
                            
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
                                    
                                        if want_nav and not want_aum:
                                            # Only carry-forward NAV (holidays/weekends have no new NAV)
                                            df_final[curr_col] = df_final[curr_col].fillna(df_final[prev_col])
                                        elif want_aum and not want_nav:
                                            # AUM: do NOT carry forward — each date gets its own Fallback_AUM
                                            pass
                                        else:
                                            # Carry forward NAV only; AUM is already populated per-day via Fallback_AUM
                                            df_final[f"{curr_col} (NAV)"] = df_final[f"{curr_col} (NAV)"].fillna(df_final[f"{prev_col} (NAV)"])
                                            # Do NOT carry forward AUM or Flows — they vary per day
                                        
                                # Fill any remaining NaNs in AUM columns with fallback values
                                if want_aum:
                                    df_pivot_fallback = df_raw.pivot(index="Scheme Code", columns="Date", values="Fallback_AUM").reset_index()
                                    fallback_cols_map = {}
                                    for d in target_dates:
                                        if want_aum and not want_nav:
                                            fallback_cols_map[d] = f"{d}_fallback_temp"
                                        elif want_nav and want_aum:
                                            fallback_cols_map[f"{d} (AUM)"] = f"{d}_fallback_temp"
                                        
                                    if fallback_cols_map:
                                        df_pivot_fallback_renamed = df_pivot_fallback.rename(columns={d: f"{d}_fallback_temp" for d in target_dates if d in df_pivot_fallback.columns})
                                        available_temp_cols = [col for col in fallback_cols_map.values() if col in df_pivot_fallback_renamed.columns]
                                        df_final_temp = pd.merge(df_final, df_pivot_fallback_renamed[["Scheme Code"] + available_temp_cols], on="Scheme Code", how="left")
                                        for main_col, temp_col in fallback_cols_map.items():
                                            if main_col in df_final.columns and temp_col in df_final_temp.columns:
                                                df_final[main_col] = df_final[main_col].fillna(df_final_temp[temp_col])
                                    
                                # Convert df_final to vertical layout
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
                                        if want_nav and not want_aum:
                                            r_item["NAV Date"] = d
                                            r_item["NAV"] = row[d]
                                        elif want_aum and not want_nav:
                                            r_item["AUM Date"] = d
                                            r_item["AUM"] = row[d]
                                        else:
                                            r_item["NAV Date"] = d
                                            r_item["NAV"] = row[f"{d} (NAV)"]
                                            r_item["AUM Date"] = d
                                            r_item["AUM"] = row[f"{d} (AUM)"]
                                        vertical_rows.append(r_item)

                                ordered_cols = [
                                    "Asset Class", 
                                    "Scheme Code", 
                                    "ISIN Div Payout / ISIN Growth", 
                                    "ISIN Div Reinvestment", 
                                    "Scheme Name", 
                                    "Plan Type", 
                                    "Option Type"
                                ]
                                if want_nav:
                                    ordered_cols.extend(["NAV Date", "NAV"])
                                if want_aum:
                                    ordered_cols.extend(["AUM Date", "AUM"])
                            
                                if vertical_rows:
                                    df_final = pd.DataFrame(vertical_rows)
                                    if "NAV Date" in df_final.columns:
                                        parsed = parse_amfi_date_series(df_final["NAV Date"])
                                        df_final["NAV Date"] = parsed.dt.strftime("%d-%m-%Y")
                                    if "AUM Date" in df_final.columns:
                                        parsed = parse_amfi_date_series(df_final["AUM Date"])
                                        df_final["AUM Date"] = parsed.dt.strftime("%d-%m-%Y")
                                    df_final = df_final[ordered_cols]
                                    df_final = df_final.sort_values(by=["Asset Class", "Scheme Name"]).reset_index(drop=True)
                                    if want_flows:
                                        df_final = calculate_flows_for_dataframe(df_final, start_date, ["Asset Class", "Scheme Code", "ISIN Div Payout / ISIN Growth", "ISIN Div Reinvestment", "Scheme Name", "Plan Type", "Option Type"])
                                        is_aum_only = False
                                else:
                                    df_final = pd.DataFrame(columns=ordered_cols)
                            
                                st.success(f"Successfully processed {len(df_final)} vertical records!")
                            
                                render_section_header("👁️", "Data Preview")
                                st.dataframe(
                                    df_final,
                                    use_container_width=True,
                                    column_config={
                                        "Daily return": st.column_config.NumberColumn(
                                            "Daily return",
                                            format="%.2f%%"
                                        )
                                    }
                                )
                            
                                excel_bytes = generate_historical_excel(df_final, [], is_aum_only=is_aum_only)
                            
                                st.download_button(
                                    label="📥 Download Styled Excel (.xlsx)",
                                    data=excel_bytes,
                                    file_name=f"amfi_nav_export_{start_date}_to_{end_date}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                            
                        except Exception as e:
                            st.error(f"Failed to fetch or process data: {e}")


    if selected_summary is not None and not selected_rows.empty:
        with finance_panel("Fund Analysis"):
            render_result_card(selected_summary, selected_rows)
            render_section_header("📑", "Latest NAV Table")
            latest_rows = selected_rows[["Scheme Name", "Scheme Code", "AMC Name", "Plan Type", "Option Type", "NAV", "NAV Date"]].copy()
            latest_rows["NAV Date"] = pd.to_datetime(latest_rows["NAV Date"], errors="coerce").dt.strftime("%d-%m-%Y")
            st.dataframe(latest_rows, use_container_width=True, hide_index=True)

            export_frame = build_export_frame(selected_rows)
            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    "Download CSV",
                    data=frame_to_csv_bytes(export_frame),
                    file_name="mutual_fund_nav_results.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
            with c2:
                st.download_button(
                    "Download Excel",
                    data=frame_to_excel_bytes(export_frame),
                    file_name="mutual_fund_nav_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            render_history(selected_summary)
            render_sip_calculator(float(selected_summary.get("Latest NAV", 0)) if pd.notna(selected_summary.get("Latest NAV")) else None)

    elif search_mode == "Batch Search" and not batch_export_rows.empty:
        export_frame = build_export_frame(batch_export_rows)
        st.download_button(
            "Download batch CSV",
            data=frame_to_csv_bytes(export_frame),
            file_name="mutual_fund_batch_results.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with finance_panel("Market Snapshot"):
        with st.expander("Browse latest market snapshot", expanded=False):
            summary = summarize_families(nav_data)
            if summary.empty:
                st.info("No fund summaries are available.")
            else:
                browse = summary[["Family Name", "AMC Name", "Latest Scheme Name", "Latest Scheme Code", "Latest NAV", "Latest NAV Date", "Plan Types", "Option Types"]].copy()
                browse["Latest NAV Date"] = pd.to_datetime(browse["Latest NAV Date"], errors="coerce").dt.strftime("%d-%m-%Y")
                st.dataframe(browse.head(200), use_container_width=True, hide_index=True)

        render_section_header("⚡", "Quick Export", "Jump straight to Excel by fund name or Scheme Code")
        quick_query = st.text_input(
            "Enter full/partial fund name or exact Scheme Code",
            placeholder="e.g. HDFC Equity or 120439",
            label_visibility="collapsed",
        )
        if quick_query.strip():
            if st.button("Export NAV Excel", key="quick_export"):
                # Try direct Scheme Code match first
                q = quick_query.strip()
                found_family_key = None
                # exact numeric or string match against Scheme Code
                exact_code = nav_data[nav_data["Scheme Code"].astype(str).str.strip().eq(q)]
                if not exact_code.empty:
                    found_family_key = exact_code.iloc[0]["Family Key"]
                else:
                    # fuzzy search across families
                    matches = search_fund(nav_data, q, plan_filter=None, option_filter=None, limit=1)
                    if not matches.empty:
                        found_family_key = matches.iloc[0]["Family Key"]

                if not found_family_key:
                    st.warning("No matching fund found for export. Try a different name or the exact Scheme Code.")
                else:
                    latest_rows = filter_family_rows(nav_data, found_family_key)
                    history = load_historical_snapshots(found_family_key)

                    if latest_rows.empty and (history is None or history.empty):
                        st.info("No NAV rows available for the selected fund.")
                    else:
                        # Build Excel with two sheets when history exists
                        out = BytesIO()
                        with pd.ExcelWriter(out, engine="openpyxl") as writer:
                            if not latest_rows.empty:
                                export_latest = latest_rows.copy()
                                if "NAV Date" in export_latest.columns:
                                    export_latest["NAV Date"] = pd.to_datetime(export_latest["NAV Date"], errors="coerce").dt.strftime("%d-%m-%Y")
                                export_latest.to_excel(writer, index=False, sheet_name="Latest NAVs")
                            if history is not None and not history.empty:
                                export_hist = history.copy()
                                if "NAV Date" in export_hist.columns:
                                    export_hist["NAV Date"] = pd.to_datetime(export_hist["NAV Date"], errors="coerce").dt.strftime("%d-%m-%Y")
                                export_hist.to_excel(writer, index=False, sheet_name="History")
                        out.seek(0)
                        file_name = f"nav_export_{normalize_filename(quick_query)}.xlsx"
                        st.download_button(
                            "Download NAV Excel",
                            data=out.getvalue(),
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

    render_app_footer()


if __name__ == "__main__":
    main()
