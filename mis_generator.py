"""
MIS Generator Module for Mutual Fund Portfolios
Generate professional MIS reports (MIS 1, MIS 2, MIS 3) with dynamic NAV fetching,
benchmark comparisons (Nifty 50 and individual benchmarks), and styled Excel/PDF exports.
"""

from __future__ import annotations

import io
import re
from datetime import datetime, date, timedelta
from typing import Dict, List, Tuple, Any, Optional

import pandas as pd
import numpy as np
import requests
import streamlit as st

# openpyxl styling imports
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ReportLab PDF imports
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Re-use existing NAV fetch services from parent codebase
from nav_fetcher import (
    fetch_amfi_data_chunked,
    fetch_nifty_data_series,
    fetch_latest_navs,
    parse_amfi_date_series,
    build_date_cols,
    _parse_amfi_date_str,
)
from ui_theme import render_section_header, render_info_card, finance_panel

# ─── Default Sample Portfolio ──────────────────────────────────────────────────

DEFAULT_SAMPLE_PORTFOLIO = [
    {
        "Scheme Name": "HDFC Flexi Cap Fund - Direct Plan - Growth",
        "ISIN": "INF179K01BE2",
        "Allocation (%)": 25.0,
        "Benchmark": "NIFTY 500",
    },
    {
        "Scheme Name": "SBI Bluechip Fund - Direct Plan - Growth",
        "ISIN": "INF200K01VT2",
        "Allocation (%)": 25.0,
        "Benchmark": "NIFTY 50",
    },
    {
        "Scheme Name": "Nippon India Small Cap Fund - Direct Plan - Growth",
        "ISIN": "INF204K01HY2",
        "Allocation (%)": 25.0,
        "Benchmark": "NIFTY Smallcap 250",
    },
    {
        "Scheme Name": "Mirae Asset Large & Midcap Fund - Direct Plan - Growth",
        "ISIN": "INF769K01150",
        "Allocation (%)": 25.0,
        "Benchmark": "NIFTY LargeMidcap 250",
    },
]

# ─── Helper Date Functions ────────────────────────────────────────────────────

def get_fy_start_date(target_date: date) -> date:
    """Return April 1st of the financial year corresponding to target_date.
    
    If target_date is July 2026, FY Start is 1-Apr-2026.
    If target_date is Feb 2026, FY Start is 1-Apr-2025.
    """
    if target_date.month >= 4:
        return date(target_date.year, 4, 1)
    else:
        return date(target_date.year - 1, 4, 1)


def get_ytd_start_date(target_date: date) -> date:
    """Return Dec 31 of the previous calendar year relative to target_date."""
    return date(target_date.year - 1, 12, 31)


def get_mtd_start_date(target_date: date) -> date:
    """Return the last day of the previous calendar month relative to target_date."""
    first_of_curr_month = date(target_date.year, target_date.month, 1)
    return first_of_curr_month - timedelta(days=1)


def get_monthly_baseline_date(target_date: date) -> date:
    """Return 1 month prior to target_date."""
    return target_date - timedelta(days=30)


def get_daily_baseline_date(target_date: date) -> date:
    """Return 1 day prior to target_date (baseline will fetch closest prior trading day)."""
    return target_date - timedelta(days=1)


# ─── Data Ingestion & Validation ──────────────────────────────────────────────

def validate_and_normalize_portfolio(df_input: pd.DataFrame) -> Tuple[pd.DataFrame, List[str], List[str]]:
    """Validate user portfolio input and prepare normalized allocations.
    
    Returns (cleaned_df, warnings, errors).
    """
    warnings: List[str] = []
    errors: List[str] = []

    if df_input.empty:
        errors.append("Portfolio input is empty. Please enter schemes manually or upload an Excel file.")
        return pd.DataFrame(), warnings, errors

    df = df_input.copy()

    # Column normalization
    isin_col = None
    name_col = None
    alloc_col = None
    bm_col = None

    for c in df.columns:
        c_low = str(c).lower().strip()
        if "isin" in c_low:
            isin_col = c
        elif "scheme" in c_low or "fund" in c_low or "name" in c_low:
            name_col = c
        elif "alloc" in c_low or "weight" in c_low or "wt" in c_low or "%" in c_low:
            alloc_col = c
        elif "bench" in c_low or "bm" in c_low:
            bm_col = c

    if not isin_col:
        errors.append("Missing required column 'ISIN' in input.")
        return pd.DataFrame(), warnings, errors

    clean_rows = []
    seen_isins = set()

    for idx, row in df.iterrows():
        isin = str(row.get(isin_col, "")).strip().upper()
        if not isin or isin in ("NAN", "NONE", "-"):
            continue

        if isin in seen_isins:
            warnings.append(f"Duplicate ISIN '{isin}' detected. Only the first occurrence was retained.")
            continue
        seen_isins.add(isin)

        name = str(row.get(name_col, isin)).strip() if name_col else isin
        if not name or name in ("NAN", "NONE"):
            name = isin

        alloc_raw = row.get(alloc_col, 0.0) if alloc_col else 0.0
        try:
            alloc_str = str(alloc_raw).replace("%", "").strip()
            alloc_val = float(alloc_str)
        except Exception:
            alloc_val = 0.0

        bm_val = str(row.get(bm_col, "NIFTY 50")).strip() if bm_col else "NIFTY 50"
        if not bm_val or bm_val in ("NAN", "NONE", "-"):
            bm_val = "NIFTY 50"

        clean_rows.append({
            "Scheme Name": name,
            "ISIN": isin,
            "Allocation (%)": alloc_val,
            "Benchmark": bm_val,
        })

    if not clean_rows:
        errors.append("No valid ISIN rows found in the portfolio input.")
        return pd.DataFrame(), warnings, errors

    res_df = pd.DataFrame(clean_rows)

    # Check total allocation
    tot_alloc = res_df["Allocation (%)"].sum()
    if tot_alloc <= 0:
        warnings.append("Total allocation was 0%. Weights have been distributed equally.")
        res_df["Allocation (%)"] = 100.0 / len(res_df)
        tot_alloc = 100.0

    if abs(tot_alloc - 100.0) > 0.01:
        warnings.append(
            f"Total allocation sum is {tot_alloc:.2f}% (not 100%). Weights have been automatically normalized."
        )

    res_df["Normalized_Weight"] = res_df["Allocation (%)"] / tot_alloc
    return res_df, warnings, errors


# ─── NAV & Benchmark Fetching Engine ──────────────────────────────────────────

def fetch_mis_nav_history(
    isin_list: List[str],
    earliest_date: date,
    end_date: date,
) -> pd.DataFrame:
    """Fetch complete historical NAVs for all ISINs from earliest_date to end_date.
    
    Includes live feed overlay for today's NAV if applicable.
    """
    # Fetch 15 days earlier to cover holiday/weekend lookups
    fetch_start = earliest_date - timedelta(days=15)
    df_raw = fetch_amfi_data_chunked(fetch_start, end_date, isin_list)
    if df_raw.empty:
        return pd.DataFrame()

    df_raw["NAV Date"] = parse_amfi_date_series(df_raw["NAV Date"])
    df_raw = df_raw.dropna(subset=["NAV Date", "NAV"])
    df_raw["ISIN_G"] = df_raw["ISIN Div Payout / ISIN Growth"].astype(str).str.strip().str.upper()
    df_raw["ISIN_R"] = df_raw["ISIN Div Reinvestment"].astype(str).str.strip().str.upper()
    df_raw["NAV Date_Date"] = df_raw["NAV Date"].dt.date

    # Overlay live feed
    latest_live = fetch_latest_navs(isin_list)
    if latest_live:
        live_rows = []
        for isin_u, info in latest_live.items():
            parsed_iso = _parse_amfi_date_str(info["date"])
            if parsed_iso:
                l_dt = datetime.strptime(parsed_iso, "%Y-%m-%d").date()
                live_rows.append({
                    "ISIN_G": isin_u,
                    "ISIN_R": isin_u,
                    "Scheme Name": info.get("scheme_name", isin_u),
                    "NAV": info["nav"],
                    "NAV Date_Date": l_dt,
                })
        if live_rows:
            df_live = pd.DataFrame(live_rows)
            df_raw = pd.concat([df_raw, df_live], ignore_index=True)

    return df_raw


def get_asof_nav(
    df_nav_raw: pd.DataFrame,
    isin: str,
    target_date: date
) -> Optional[float]:
    """Find the latest available NAV for an ISIN on or before target_date."""
    isin_u = isin.strip().upper()
    mask = (
        (df_nav_raw["ISIN_G"] == isin_u) | (df_nav_raw["ISIN_R"] == isin_u)
    ) & (df_nav_raw["NAV Date_Date"] <= target_date)

    matched = df_nav_raw[mask].sort_values("NAV Date_Date", ascending=False)
    if not matched.empty:
        val = matched.iloc[0]["NAV"]
        if pd.notna(val) and float(val) > 0:
            return float(val)
    return None


def fetch_benchmark_series(
    benchmark_names: List[str],
    start_date: date,
    end_date: date
) -> Dict[str, Dict[date, float]]:
    """Fetch daily price series for all required benchmark indices.
    
    Returns dict: {benchmark_name: {date_obj: close_price}}
    """
    fetch_start = start_date - timedelta(days=15)
    benchmarks_data = {}

    for bm in set(benchmark_names):
        bm_clean = bm.strip() if bm and str(bm).strip() else "NIFTY 50"
        raw_map = fetch_nifty_data_series(bm_clean, fetch_start, end_date)
        date_obj_map = {}
        for k_str, val in raw_map.items():
            try:
                dt_o = datetime.strptime(k_str, "%d-%m-%Y").date()
                date_obj_map[dt_o] = val
            except Exception:
                pass
        benchmarks_data[bm_clean] = date_obj_map

    return benchmarks_data


def get_asof_benchmark_price(
    bm_map: Dict[date, float],
    target_date: date
) -> Optional[float]:
    """Find the latest available benchmark price on or before target_date."""
    if not bm_map:
        return None
    valid_dates = [d for d in bm_map.keys() if d <= target_date]
    if valid_dates:
        latest_d = max(valid_dates)
        return bm_map[latest_d]
    return None


# ─── Return & MIS Core Calculations ─────────────────────────────────────────

def calc_return(end_val: Optional[float], start_val: Optional[float]) -> Optional[float]:
    """Calculate non-annualized return %: ((end - start) / start) * 100."""
    if end_val is not None and start_val is not None and start_val > 0:
        return ((end_val - start_val) / start_val) * 100.0
    return None


def generate_mis_reports_data(
    portfolio_df: pd.DataFrame,
    start_date: date,
    end_date: date
) -> Dict[str, Any]:
    """Calculate all returns, excess returns, and weighted portfolio metrics for MIS 1, 2 & 3.
    
    Returns a dictionary containing MIS dataframes and executive summary statistics.
    """
    # 1. Determine key baseline dates
    d_end = end_date
    d_start = start_date
    d_daily_base = get_daily_baseline_date(end_date)
    d_monthly_base = get_monthly_baseline_date(end_date)
    d_mtd_base = get_mtd_start_date(end_date)
    d_ytd_base = get_ytd_start_date(end_date)
    d_fy_base = get_fy_start_date(end_date)

    earliest_date = min(d_start, d_daily_base, d_monthly_base, d_mtd_base, d_ytd_base, d_fy_base)

    # 2. Fetch raw NAVs and Benchmark series
    isins = portfolio_df["ISIN"].tolist()
    df_nav_raw = fetch_mis_nav_history(isins, earliest_date, end_date)

    if df_nav_raw.empty:
        raise ValueError("Could not fetch historical NAV data for the entered portfolio ISINs.")

    bm_list = portfolio_df["Benchmark"].tolist() + ["NIFTY 50"]
    benchmarks_data = fetch_benchmark_series(bm_list, earliest_date, end_date)
    nifty50_map = benchmarks_data.get("NIFTY 50", {})

    # Nifty 50 baseline prices
    nifty_nav_end = get_asof_benchmark_price(nifty50_map, d_end)
    nifty_nav_daily_base = get_asof_benchmark_price(nifty50_map, d_daily_base)
    nifty_nav_period_base = get_asof_benchmark_price(nifty50_map, d_start)
    nifty_nav_fy_base = get_asof_benchmark_price(nifty50_map, d_fy_base)

    nifty_daily_ret = calc_return(nifty_nav_end, nifty_nav_daily_base)
    nifty_period_ret = calc_return(nifty_nav_end, nifty_nav_period_base)
    nifty_fy_ret = calc_return(nifty_nav_end, nifty_nav_fy_base)

    # 3. Calculate per-scheme metrics
    mis1_rows = []
    mis2_rows = []
    mis3_rows = []

    for _, row in portfolio_df.iterrows():
        name = row["Scheme Name"]
        isin = row["ISIN"]
        alloc = row["Allocation (%)"]
        weight = row["Normalized_Weight"]
        bm_name = row["Benchmark"]

        # As-of NAVs for scheme
        nav_end = get_asof_nav(df_nav_raw, isin, d_end)
        nav_daily_base = get_asof_nav(df_nav_raw, isin, d_daily_base)
        nav_period_base = get_asof_nav(df_nav_raw, isin, d_start)
        nav_fy_base = get_asof_nav(df_nav_raw, isin, d_fy_base)

        # Scheme returns
        sch_daily_ret = calc_return(nav_end, nav_daily_base)
        sch_period_ret = calc_return(nav_end, nav_period_base)
        sch_fy_ret = calc_return(nav_end, nav_fy_base)

        # Benchmark prices & returns
        bm_map = benchmarks_data.get(bm_name, nifty50_map)
        bm_nav_end = get_asof_benchmark_price(bm_map, d_end)
        bm_nav_period_base = get_asof_benchmark_price(bm_map, d_start)
        bm_nav_fy_base = get_asof_benchmark_price(bm_map, d_fy_base)

        bm_period_ret = calc_return(bm_nav_end, bm_nav_period_base)
        bm_fy_ret = calc_return(bm_nav_end, bm_nav_fy_base)

        # --- MIS 1 (Portfolio vs Nifty) ---
        daily_excess_nifty = (sch_daily_ret - nifty_daily_ret) if (sch_daily_ret is not None and nifty_daily_ret is not None) else None
        period_excess_nifty = (sch_period_ret - nifty_period_ret) if (sch_period_ret is not None and nifty_period_ret is not None) else None

        mis1_rows.append({
            "Scheme Name": name,
            "ISIN": isin,
            "Allocation (%)": alloc,
            "Normalized_Weight": weight,
            "Daily Return (%)": sch_daily_ret,
            "Nifty Daily Return (%)": nifty_daily_ret,
            "Daily Excess (%)": daily_excess_nifty,
            "Period Return (%)": sch_period_ret,
            "Nifty Period Return (%)": nifty_period_ret,
            "Period Excess (%)": period_excess_nifty,
        })

        # --- MIS 2 (Portfolio vs Individual Benchmark) ---
        period_excess_bm = (sch_period_ret - bm_period_ret) if (sch_period_ret is not None and bm_period_ret is not None) else None

        mis2_rows.append({
            "Scheme Name": name,
            "ISIN": isin,
            "Allocation (%)": alloc,
            "Normalized_Weight": weight,
            "Benchmark": bm_name,
            "Period Return (%)": sch_period_ret,
            "Benchmark Return (%)": bm_period_ret,
            "Excess Return (%)": period_excess_bm,
        })

        # --- MIS 3 (Performance Since FY Start - 1 April) ---
        excess_over_bm_fy = (sch_fy_ret - bm_fy_ret) if (sch_fy_ret is not None and bm_fy_ret is not None) else None
        excess_over_nifty_fy = (sch_fy_ret - nifty_fy_ret) if (sch_fy_ret is not None and nifty_fy_ret is not None) else None

        mis3_rows.append({
            "Scheme Name": name,
            "ISIN": isin,
            "Allocation (%)": alloc,
            "Normalized_Weight": weight,
            "Scheme Return (%)": sch_fy_ret,
            "Benchmark Return (%)": bm_fy_ret,
            "Nifty Return (%)": nifty_fy_ret,
            "Excess over Benchmark (%)": excess_over_bm_fy,
            "Excess over Nifty (%)": excess_over_nifty_fy,
        })

    df_mis1 = pd.DataFrame(mis1_rows)
    df_mis2 = pd.DataFrame(mis2_rows)
    df_mis3 = pd.DataFrame(mis3_rows)

    # 4. Weighted Portfolio Totals
    def weighted_sum(series: pd.Series, weights: pd.Series) -> Optional[float]:
        valid_mask = series.notna() & weights.notna()
        if not valid_mask.any():
            return None
        w_sub = weights[valid_mask]
        s_sub = series[valid_mask]
        norm_w = w_sub / w_sub.sum()
        return (s_sub * norm_w).sum()

    # MIS 1 Summary Totals
    w_sch_daily = weighted_sum(df_mis1["Daily Return (%)"], df_mis1["Normalized_Weight"])
    w_nifty_daily = nifty_daily_ret
    w_daily_excess = (w_sch_daily - w_nifty_daily) if (w_sch_daily is not None and w_nifty_daily is not None) else None

    w_sch_period = weighted_sum(df_mis1["Period Return (%)"], df_mis1["Normalized_Weight"])
    w_nifty_period = nifty_period_ret
    w_period_excess = (w_sch_period - w_nifty_period) if (w_sch_period is not None and w_nifty_period is not None) else None

    mis1_summary = {
        "Scheme Name": "PORTFOLIO WEIGHTED TOTAL",
        "ISIN": "-",
        "Allocation (%)": df_mis1["Allocation (%)"].sum(),
        "Daily Return (%)": w_sch_daily,
        "Nifty Daily Return (%)": w_nifty_daily,
        "Daily Excess (%)": w_daily_excess,
        "Period Return (%)": w_sch_period,
        "Nifty Period Return (%)": w_nifty_period,
        "Period Excess (%)": w_period_excess,
    }

    # MIS 2 Summary Totals
    w_bm_period = weighted_sum(df_mis2["Benchmark Return (%)"], df_mis2["Normalized_Weight"])
    w_excess_bm = (w_sch_period - w_bm_period) if (w_sch_period is not None and w_bm_period is not None) else None

    mis2_summary = {
        "Scheme Name": "PORTFOLIO WEIGHTED TOTAL",
        "ISIN": "-",
        "Allocation (%)": df_mis2["Allocation (%)"].sum(),
        "Benchmark": "Weighted Benchmarks",
        "Period Return (%)": w_sch_period,
        "Benchmark Return (%)": w_bm_period,
        "Excess Return (%)": w_excess_bm,
    }

    # MIS 3 Summary Totals
    w_sch_fy = weighted_sum(df_mis3["Scheme Return (%)"], df_mis3["Normalized_Weight"])
    w_bm_fy = weighted_sum(df_mis3["Benchmark Return (%)"], df_mis3["Normalized_Weight"])
    w_nifty_fy = nifty_fy_ret
    w_excess_bm_fy = (w_sch_fy - w_bm_fy) if (w_sch_fy is not None and w_bm_fy is not None) else None
    w_excess_nifty_fy = (w_sch_fy - w_nifty_fy) if (w_sch_fy is not None and w_nifty_fy is not None) else None

    mis3_summary = {
        "Scheme Name": "PORTFOLIO WEIGHTED TOTAL",
        "ISIN": "-",
        "Allocation (%)": df_mis3["Allocation (%)"].sum(),
        "Scheme Return (%)": w_sch_fy,
        "Benchmark Return (%)": w_bm_fy,
        "Nifty Return (%)": w_nifty_fy,
        "Excess over Benchmark (%)": w_excess_bm_fy,
        "Excess over Nifty (%)": w_excess_nifty_fy,
    }

    return {
        "mis1": df_mis1,
        "mis1_summary": mis1_summary,
        "mis2": df_mis2,
        "mis2_summary": mis2_summary,
        "mis3": df_mis3,
        "mis3_summary": mis3_summary,
        "dates": {
            "start_date": d_start,
            "end_date": d_end,
            "fy_start": d_fy_base,
        },
    }


# ─── Excel Workbook Export Engine ─────────────────────────────────────────────

def export_mis_to_excel(mis_data: Dict[str, Any]) -> bytes:
    """Generate a corporate-styled Excel workbook containing Executive Summary, MIS 1, MIS 2, and MIS 3 tabs."""
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    font_family = "Segoe UI"
    title_font = Font(name=font_family, size=14, bold=True, color="1F497D")
    meta_label_font = Font(name=font_family, size=10, bold=True, color="555555")
    meta_val_font = Font(name=font_family, size=10, color="000000")

    header_fill_dark = PatternFill("solid", fgColor="1F497D")  # Dark Blue
    header_fill_green = PatternFill("solid", fgColor="145A32") # Dark Emerald
    header_fill_gold = PatternFill("solid", fgColor="7D5A1F")  # Gold/Bronze
    summary_fill = PatternFill("solid", fgColor="D9E1F2")      # Light Blue Accent

    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    data_font = Font(name=font_family, size=10)
    summary_font = Font(name=font_family, size=10, bold=True)

    even_fill = PatternFill("solid", fgColor="F9FAFB")
    odd_fill = PatternFill("solid", fgColor="FFFFFF")

    border_thin = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    d_start = mis_data["dates"]["start_date"].strftime("%d-%m-%Y")
    d_end = mis_data["dates"]["end_date"].strftime("%d-%m-%Y")
    d_fy = mis_data["dates"]["fy_start"].strftime("%d-%m-%Y")

    def format_worksheet_table(
        ws: openpyxl.worksheet.worksheet.Worksheet,
        title_text: str,
        df: pd.DataFrame,
        summary_dict: dict,
        header_fill: PatternFill,
        num_pct_cols: List[str],
    ):
        ws.views.sheetView[0].showGridLines = True
        ws.row_dimensions[1].height = 24
        ws.row_dimensions[2].height = 18

        ws.cell(row=1, column=1, value=title_text).font = title_font
        ws.cell(row=2, column=1, value=f"Date Range: {d_start} to {d_end} | FY Start: {d_fy}").font = meta_label_font

        # Headers start at row 4
        headers = [c for c in df.columns if c != "Normalized_Weight"]
        ws.row_dimensions[4].height = 28

        for ci, col in enumerate(headers, 1):
            cell = ws.cell(row=4, column=ci, value=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border_thin
            cell.alignment = align_left if col in ("Scheme Name", "ISIN", "Benchmark") else align_right

        # Data Rows
        curr_row = 5
        for _, row in df.iterrows():
            ws.row_dimensions[curr_row].height = 20
            row_fill = even_fill if curr_row % 2 == 0 else odd_fill

            for ci, col in enumerate(headers, 1):
                val = row.get(col)
                cell = ws.cell(row=curr_row, column=ci)

                if col in ("Scheme Name", "ISIN", "Benchmark"):
                    cell.value = str(val) if pd.notna(val) else "-"
                    cell.alignment = align_left
                elif col == "Allocation (%)":
                    cell.value = (val / 100.0) if pd.notna(val) else 0.0
                    cell.number_format = "0.00%"
                    cell.alignment = align_right
                elif col in num_pct_cols:
                    if pd.notna(val):
                        cell.value = val / 100.0
                        cell.number_format = "0.00%"
                    else:
                        cell.value = "N/A"
                    cell.alignment = align_right

                cell.font = data_font
                cell.fill = row_fill
                cell.border = border_thin
            curr_row += 1

        # Summary Total Row
        ws.row_dimensions[curr_row].height = 22
        for ci, col in enumerate(headers, 1):
            val = summary_dict.get(col)
            cell = ws.cell(row=curr_row, column=ci)

            if col in ("Scheme Name", "ISIN", "Benchmark"):
                cell.value = str(val) if pd.notna(val) else "-"
                cell.alignment = align_left
            elif col == "Allocation (%)":
                cell.value = (val / 100.0) if pd.notna(val) else 0.0
                cell.number_format = "0.00%"
                cell.alignment = align_right
            elif col in num_pct_cols:
                if pd.notna(val):
                    cell.value = val / 100.0
                    cell.number_format = "0.00%"
                else:
                    cell.value = "N/A"
                cell.alignment = align_right

            cell.font = summary_font
            cell.fill = summary_fill
            cell.border = border_thin
        
        # Auto-adjust column widths
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # 1. Sheet MIS 1 - Portfolio vs Nifty
    ws1 = wb.create_sheet(title="MIS 1 - Portfolio vs Nifty")
    format_worksheet_table(
        ws1,
        "MIS 1: Portfolio vs Nifty Performance Report",
        mis_data["mis1"],
        mis_data["mis1_summary"],
        header_fill_dark,
        [
            "Daily Return (%)", "Nifty Daily Return (%)", "Daily Excess (%)",
            "Period Return (%)", "Nifty Period Return (%)", "Period Excess (%)"
        ],
    )

    # 2. Sheet MIS 2 - Portfolio vs Benchmark
    ws2 = wb.create_sheet(title="MIS 2 - vs Benchmarks")
    format_worksheet_table(
        ws2,
        "MIS 2: Portfolio vs Individual Benchmark Report",
        mis_data["mis2"],
        mis_data["mis2_summary"],
        header_fill_green,
        ["Period Return (%)", "Benchmark Return (%)", "Excess Return (%)"],
    )

    # 3. Sheet MIS 3 - FY Performance
    ws3 = wb.create_sheet(title="MIS 3 - Since FY Start")
    format_worksheet_table(
        ws3,
        "MIS 3: Performance Since Financial Year Start (1 April)",
        mis_data["mis3"],
        mis_data["mis3_summary"],
        header_fill_gold,
        [
            "Scheme Return (%)", "Benchmark Return (%)", "Nifty Return (%)",
            "Excess over Benchmark (%)", "Excess over Nifty (%)"
        ],
    )

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── PDF Report Export Engine ──────────────────────────────────────────────────

def export_mis_to_pdf(mis_data: Dict[str, Any]) -> bytes:
    """Generate a clean, professional landscape PDF report containing MIS 1, MIS 2, and MIS 3 tables."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "PDFTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=20,
        textColor=colors.HexColor("#1F497D"),
        spaceAfter=4,
    )

    sub_style = ParagraphStyle(
        "PDFSubTitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#555555"),
        spaceAfter=12,
    )

    section_header_style = ParagraphStyle(
        "PDFSectionHeader",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=15,
        textColor=colors.HexColor("#1F497D"),
        spaceBefore=8,
        spaceAfter=6,
    )

    cell_hdr_style = ParagraphStyle(
        "PDFCellHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
        textColor=colors.white,
        alignment=1, # Center
    )

    cell_data_style = ParagraphStyle(
        "PDFCellData",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9,
        alignment=0, # Left
    )

    cell_num_style = ParagraphStyle(
        "PDFCellNum",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9,
        alignment=2, # Right
    )

    cell_summary_style = ParagraphStyle(
        "PDFCellSummary",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.5,
        leading=9,
        alignment=2, # Right
    )

    d_start = mis_data["dates"]["start_date"].strftime("%d-%b-%Y")
    d_end = mis_data["dates"]["end_date"].strftime("%d-%b-%Y")
    d_fy = mis_data["dates"]["fy_start"].strftime("%d-%b-%Y")

    story = []

    # Title & Metadata
    story.append(Paragraph("MUTUAL FUND PORTFOLIO MIS REPORT", title_style))
    story.append(Paragraph(
        f"<b>Selected Date Range:</b> {d_start} to {d_end} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Financial Year Start:</b> {d_fy}",
        sub_style
    ))

    def build_pdf_table(
        df: pd.DataFrame,
        summary_dict: dict,
        num_cols: List[str],
        col_widths: List[float],
        header_bg: colors.Color,
    ) -> Table:
        headers = [c for c in df.columns if c != "Normalized_Weight"]
        table_data = []

        # Header Row
        hdr_row = [Paragraph(h, cell_hdr_style) for h in headers]
        table_data.append(hdr_row)

        # Data Rows
        for _, row in df.iterrows():
            r_cells = []
            for h in headers:
                val = row.get(h)
                if h in ("Scheme Name", "ISIN", "Benchmark"):
                    r_cells.append(Paragraph(str(val) if pd.notna(val) else "-", cell_data_style))
                elif h == "Allocation (%)":
                    v_str = f"{val:.2f}%" if pd.notna(val) else "0.00%"
                    r_cells.append(Paragraph(v_str, cell_num_style))
                elif h in num_cols:
                    v_str = f"{val:.2f}%" if pd.notna(val) else "N/A"
                    r_cells.append(Paragraph(v_str, cell_num_style))
            table_data.append(r_cells)

        # Summary Row
        sum_cells = []
        for h in headers:
            val = summary_dict.get(h)
            if h in ("Scheme Name", "ISIN", "Benchmark"):
                sum_cells.append(Paragraph(str(val), cell_summary_style))
            elif h == "Allocation (%)":
                v_str = f"{val:.2f}%" if pd.notna(val) else "0.00%"
                sum_cells.append(Paragraph(v_str, cell_summary_style))
            elif h in num_cols:
                v_str = f"{val:.2f}%" if pd.notna(val) else "N/A"
                sum_cells.append(Paragraph(v_str, cell_summary_style))
        table_data.append(sum_cells)

        t = Table(table_data, colWidths=col_widths, repeatRows=1)
        t_style = [
            ("BACKGROUND", (0, 0), (-1, 0), header_bg),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D3D3D3")),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#EAECEE")),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]
        # Alternating row colors
        for i in range(1, len(table_data) - 1):
            if i % 2 == 0:
                t_style.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F8F9F9")))

        t.setStyle(TableStyle(t_style))
        return t

    # 1. MIS 1 Table
    story.append(Paragraph("MIS 1: Portfolio vs Nifty Performance", section_header_style))
    col_w1 = [200, 85, 55, 60, 60, 60, 60, 60, 60] # Total ~ 800
    t1 = build_pdf_table(
        mis_data["mis1"],
        mis_data["mis1_summary"],
        [
            "Daily Return (%)", "Nifty Daily Return (%)", "Daily Excess (%)",
            "Period Return (%)", "Nifty Period Return (%)", "Period Excess (%)"
        ],
        col_w1,
        colors.HexColor("#1F497D"),
    )
    story.append(t1)
    story.append(Spacer(1, 14))

    # 2. MIS 2 Table
    story.append(Paragraph("MIS 2: Portfolio vs Individual Benchmarks", section_header_style))
    col_w2 = [220, 90, 65, 120, 75, 75, 75]
    t2 = build_pdf_table(
        mis_data["mis2"],
        mis_data["mis2_summary"],
        ["Period Return (%)", "Benchmark Return (%)", "Excess Return (%)"],
        col_w2,
        colors.HexColor("#145A32"),
    )
    story.append(t2)
    story.append(Spacer(1, 14))

    # 3. MIS 3 Table
    story.append(Paragraph("MIS 3: Performance Since Financial Year Start (1 April)", section_header_style))
    col_w3 = [210, 85, 65, 80, 80, 80, 85, 85]
    t3 = build_pdf_table(
        mis_data["mis3"],
        mis_data["mis3_summary"],
        [
            "Scheme Return (%)", "Benchmark Return (%)", "Nifty Return (%)",
            "Excess over Benchmark (%)", "Excess over Nifty (%)"
        ],
        col_w3,
        colors.HexColor("#7D5A1F"),
    )
    story.append(t3)

    doc.build(story)
    return buf.getvalue()


# ─── Streamlit UI View Renderer ───────────────────────────────────────────────

def render_mis_generator_page():
    """Render the full interactive MIS Generator Streamlit page."""
    render_section_header(
        "📊",
        "MIS Generator Module",
        "Create custom portfolios, calculate returns, compare against benchmarks, and export MIS reports"
    )

    render_info_card(
        "<strong>MIS Generator:</strong> Enter your mutual fund portfolio manually or upload an Excel file. "
        "The app fetches complete NAV history dynamically, calculates Daily, Period, and FY returns, "
        "compares against Nifty 50 and individual fund benchmarks, and exports corporate-styled Excel and PDF reports."
    )

    # State management
    if "portfolio_input_df" not in st.session_state:
        st.session_state["portfolio_input_df"] = pd.DataFrame(DEFAULT_SAMPLE_PORTFOLIO)

    if "mis_results" not in st.session_state:
        st.session_state["mis_results"] = None

    # Step 1: Input Mode Selection
    with finance_panel("1. Portfolio Input"):
        input_mode = st.radio(
            "Portfolio Entry Method",
            ["Manual Entry / Edit", "Upload Excel File"],
            horizontal=True,
        )

        if input_mode == "Manual Entry / Edit":
            col_b1, col_b2 = st.columns([1, 4])
            with col_b1:
                if st.button("🔄 Reset Sample Template", use_container_width=True):
                    st.session_state["portfolio_input_df"] = pd.DataFrame(DEFAULT_SAMPLE_PORTFOLIO)
                    st.session_state["mis_results"] = None
                    st.rerun()

            edited_df = st.data_editor(
                st.session_state["portfolio_input_df"],
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Scheme Name": st.column_config.TextColumn("Scheme Name", width="large", required=True),
                    "ISIN": st.column_config.TextColumn("ISIN", width="medium", required=True),
                    "Allocation (%)": st.column_config.NumberColumn("Allocation (%)", min_value=0.0, max_value=100.0, format="%.2f%%"),
                    "Benchmark": st.column_config.TextColumn("Benchmark (e.g. NIFTY 50, NIFTY 500)", width="medium"),
                },
                key="mis_manual_editor",
            )
            current_portfolio_df = edited_df

        else:
            uploaded_file = st.file_uploader(
                "Upload Portfolio Excel (.xlsx)",
                type=["xlsx", "xls"],
                help="Excel file must contain columns: Scheme Name, ISIN, Allocation (%), Benchmark"
            )
            if uploaded_file is not None:
                try:
                    df_up = pd.read_excel(uploaded_file)
                    st.success(f"Successfully loaded {len(df_up)} rows from uploaded file.")
                    current_portfolio_df = df_up
                except Exception as exc:
                    st.error(f"Error reading uploaded Excel file: {exc}")
                    current_portfolio_df = st.session_state["portfolio_input_df"]
            else:
                st.info("Upload an Excel file above, or switch to Manual Entry.")
                current_portfolio_df = st.session_state["portfolio_input_df"]

        # Validate input
        clean_df, warnings, errors = validate_and_normalize_portfolio(current_portfolio_df)

        for w in warnings:
            st.warning(w)

        for e in errors:
            st.error(e)

        if not clean_df.empty:
            tot_alloc = clean_df["Allocation (%)"].sum()
            st.caption(
                f"📋 **Valid Schemes Indexed:** {len(clean_df)} | **Total Input Allocation:** {tot_alloc:.2f}% "
                f"*(weights automatically normalized for return calculations)*"
            )

    # Step 2: Date Selection & MIS Generation Trigger
    with finance_panel("2. Date Range & Generate"):
        col_d1, col_d2, col_d3 = st.columns([2, 2, 2])
        today = date.today()
        default_start = today - timedelta(days=365)

        with col_d1:
            start_date = st.date_input("Start Date", value=default_start, max_value=today)
        with col_d2:
            end_date = st.date_input("End Date", value=today, max_value=today)
        with col_d3:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            generate_btn = st.button("⚡ Generate MIS Reports", type="primary", use_container_width=True)

        if start_date > end_date:
            st.error("Start Date must be before or equal to End Date.")

        if generate_btn:
            if clean_df.empty:
                st.error("Cannot generate MIS reports without valid portfolio input.")
            elif start_date > end_date:
                st.error("Invalid date range.")
            else:
                with st.spinner("Fetching historical NAVs & generating MIS reports..."):
                    try:
                        results = generate_mis_reports_data(clean_df, start_date, end_date)
                        st.session_state["mis_results"] = results
                        st.success("MIS Reports generated successfully!")
                    except Exception as exc:
                        st.error(f"Failed to generate MIS reports: {exc}")

    # Step 3: Display MIS Reports & Export Options
    res = st.session_state.get("mis_results")
    if res:
        with finance_panel("3. MIS Reports Output"):
            tab1, tab2, tab3 = st.tabs([
                "📊 MIS 1: Portfolio vs Nifty",
                "🎯 MIS 2: Portfolio vs Benchmark",
                "📅 MIS 3: FY Performance (Since 1 Apr)",
            ])

            col_config_pct = {
                "Allocation (%)": st.column_config.NumberColumn("Allocation (%)", format="%.2f%%"),
                "Daily Return (%)": st.column_config.NumberColumn("Daily Return (%)", format="%.2f%%"),
                "Nifty Daily Return (%)": st.column_config.NumberColumn("Nifty Daily Return (%)", format="%.2f%%"),
                "Daily Excess (%)": st.column_config.NumberColumn("Daily Excess (%)", format="%.2f%%"),
                "Period Return (%)": st.column_config.NumberColumn("Period Return (%)", format="%.2f%%"),
                "Nifty Period Return (%)": st.column_config.NumberColumn("Nifty Period Return (%)", format="%.2f%%"),
                "Period Excess (%)": st.column_config.NumberColumn("Period Excess (%)", format="%.2f%%"),
                "Benchmark Return (%)": st.column_config.NumberColumn("Benchmark Return (%)", format="%.2f%%"),
                "Excess Return (%)": st.column_config.NumberColumn("Excess Return (%)", format="%.2f%%"),
                "Scheme Return (%)": st.column_config.NumberColumn("Scheme Return (%)", format="%.2f%%"),
                "Nifty Return (%)": st.column_config.NumberColumn("Nifty Return (%)", format="%.2f%%"),
                "Excess over Benchmark (%)": st.column_config.NumberColumn("Excess over Benchmark (%)", format="%.2f%%"),
                "Excess over Nifty (%)": st.column_config.NumberColumn("Excess over Nifty (%)", format="%.2f%%"),
            }

            with tab1:
                render_section_header("📈", "MIS 1: Portfolio vs Nifty Performance")
                mis1_display = pd.concat([
                    res["mis1"].drop(columns=["Normalized_Weight"], errors="ignore"),
                    pd.DataFrame([res["mis1_summary"]])
                ], ignore_index=True)
                st.dataframe(mis1_display, use_container_width=True, hide_index=True, column_config=col_config_pct)

            with tab2:
                render_section_header("🎯", "MIS 2: Portfolio vs Individual Benchmarks")
                mis2_display = pd.concat([
                    res["mis2"].drop(columns=["Normalized_Weight"], errors="ignore"),
                    pd.DataFrame([res["mis2_summary"]])
                ], ignore_index=True)
                st.dataframe(mis2_display, use_container_width=True, hide_index=True, column_config=col_config_pct)

            with tab3:
                render_section_header("📅", "MIS 3: Performance Since Financial Year Start (1 April)")
                mis3_display = pd.concat([
                    res["mis3"].drop(columns=["Normalized_Weight"], errors="ignore"),
                    pd.DataFrame([res["mis3_summary"]])
                ], ignore_index=True)
                st.dataframe(mis3_display, use_container_width=True, hide_index=True, column_config=col_config_pct)

            # Export Section
            render_section_header("📥", "Export Reports")
            col_exp1, col_exp2 = st.columns(2)

            with col_exp1:
                excel_bytes = export_mis_to_excel(res)
                st.download_button(
                    label="Download Excel MIS Report (.xlsx)",
                    data=excel_bytes,
                    file_name=f"MIS_Report_{res['dates']['start_date']}_to_{res['dates']['end_date']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            with col_exp2:
                pdf_bytes = export_mis_to_pdf(res)
                st.download_button(
                    label="Download PDF MIS Report (.pdf)",
                    data=pdf_bytes,
                    file_name=f"MIS_Report_{res['dates']['start_date']}_to_{res['dates']['end_date']}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
