"""
AMFI NAV Fetcher — Standalone UI
Fetch historical NAV data from AMFI India by ISIN(s) and/or date range.

Run with:
    streamlit run nav_fetcher.py
"""

from __future__ import annotations

import re
import sys
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import List

import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ─── Project helpers ─────────────────────────────────────────────────────────
sys.path.append(str(Path(__file__).parent))
from amfi_nav import classify_option_type, classify_plan_type

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AMFI NAV Fetcher",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Dark background */
.stApp { background: linear-gradient(135deg, #0f1724 0%, #162032 50%, #0f1f2e 100%); }

/* Hero section */
.hero-card {
    background: linear-gradient(135deg, #1a2d45 0%, #1e3a56 100%);
    border: 1px solid rgba(64, 156, 255, 0.2);
    border-radius: 16px;
    padding: 2rem 2.5rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 8px 32px rgba(0,0,0,0.4);
}
.hero-title {
    font-size: 2rem; font-weight: 700;
    background: linear-gradient(90deg, #60a5fa, #38bdf8, #34d399);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    margin: 0; padding-bottom: 0.3rem;
}
.hero-sub { color: #94a3b8; font-size: 1rem; margin: 0; }

/* Mode card */
.mode-info {
    background: rgba(30, 58, 86, 0.6);
    border: 1px solid rgba(64, 156, 255, 0.15);
    border-left: 4px solid #3b82f6;
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: 1rem;
    color: #94a3b8;
    font-size: 0.88rem;
    line-height: 1.6;
}

/* Stat chips */
.stat-row { display: flex; gap: 12px; margin-bottom: 1rem; flex-wrap: wrap; }
.stat-chip {
    background: rgba(59, 130, 246, 0.12);
    border: 1px solid rgba(59, 130, 246, 0.3);
    border-radius: 20px;
    padding: 6px 14px;
    color: #60a5fa;
    font-size: 0.82rem;
    font-weight: 500;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #2563eb, #1d4ed8) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    padding: 0.6rem 1.5rem !important;
    transition: all 0.2s !important;
    box-shadow: 0 4px 14px rgba(37,99,235,0.4) !important;
}
.stButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(37,99,235,0.55) !important;
}

/* Download button */
.stDownloadButton > button {
    background: linear-gradient(135deg, #059669, #047857) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    box-shadow: 0 4px 14px rgba(5,150,105,0.4) !important;
}
.stDownloadButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(5,150,105,0.55) !important;
}

/* Inputs */
.stTextArea textarea, .stTextInput input {
    background: rgba(15, 30, 50, 0.8) !important;
    border: 1px solid rgba(64, 156, 255, 0.25) !important;
    border-radius: 8px !important;
    color: #e2e8f0 !important;
    font-family: 'Inter', monospace !important;
}
.stTextArea textarea:focus, .stTextInput input:focus {
    border-color: #3b82f6 !important;
    box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.2) !important;
}

/* Date inputs */
.stDateInput input {
    background: rgba(15, 30, 50, 0.8) !important;
    border: 1px solid rgba(64, 156, 255, 0.25) !important;
    border-radius: 8px !important;
    color: #e2e8f0 !important;
}

/* Radio */
.stRadio > div { gap: 8px; }
.stRadio label {
    background: rgba(30, 58, 86, 0.5) !important;
    border: 1px solid rgba(64, 156, 255, 0.2) !important;
    border-radius: 8px !important;
    padding: 6px 16px !important;
    color: #94a3b8 !important;
    font-weight: 500 !important;
    transition: all 0.15s !important;
}
.stRadio label:has(input:checked) {
    background: rgba(37, 99, 235, 0.25) !important;
    border-color: #3b82f6 !important;
    color: #60a5fa !important;
}

/* Checkboxes */
.stCheckbox label { color: #94a3b8 !important; font-size: 0.9rem !important; }

/* Dataframe */
.stDataFrame { border: 1px solid rgba(64,156,255,0.15) !important; border-radius: 10px !important; overflow: hidden; }

/* Alerts */
.stSuccess { background: rgba(5,150,105,0.12) !important; border: 1px solid rgba(5,150,105,0.3) !important; border-radius: 8px !important; }
.stWarning { background: rgba(245,158,11,0.12) !important; border: 1px solid rgba(245,158,11,0.3) !important; border-radius: 8px !important; }
.stError { background: rgba(239,68,68,0.12) !important; border: 1px solid rgba(239,68,68,0.3) !important; border-radius: 8px !important; }
.stInfo { background: rgba(59,130,246,0.12) !important; border: 1px solid rgba(59,130,246,0.3) !important; border-radius: 8px !important; }

/* Labels */
label { color: #94a3b8 !important; font-weight: 500 !important; font-size: 0.88rem !important; }
h3, h4 { color: #e2e8f0 !important; }

/* Divider */
hr { border-color: rgba(64,156,255,0.1) !important; }

/* Spinner */
.stSpinner > div { color: #60a5fa !important; }
</style>
""",
    unsafe_allow_html=True,
)

# ─── Constants ────────────────────────────────────────────────────────────────
AMFI_HISTORY_URL = (
    "https://portal.amfiindia.com/DownloadNAVHistoryReport_Po.aspx"
    "?frmdt={frmdt}&todt={todt}"
)
AMFI_LATEST_URL = "https://portal.amfiindia.com/spages/NAVAll.txt"

# ─── Helpers ─────────────────────────────────────────────────────────────────

def parse_amfi_text(text: str) -> pd.DataFrame:
    """Parse the semicolon-delimited AMFI NAV feed into a DataFrame."""
    rows: List[dict] = []
    current_section = "Unknown"

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

        parts = [p.strip() for p in line.split(";")]
        if len(parts) < 8:
            continue

        scheme_code = parts[0]
        scheme_name = parts[1]
        isin_growth = parts[2] if parts[2] not in ("", "-") else None
        isin_reinvest = parts[3] if parts[3] not in ("", "-") else None
        nav_value = parts[4]
        nav_date = parts[7]

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


def style_excel(df: pd.DataFrame, date_cols: List[str]) -> bytes:
    """Write df to a styled Excel workbook and return as bytes."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="NAV Data")
        wb = writer.book
        ws = writer.sheets["NAV Data"]

        font_name = "Segoe UI"
        h_fill = PatternFill("solid", fgColor="1F497D")
        h_font = Font(name=font_name, size=11, bold=True, color="FFFFFF")
        data_font = Font(name=font_name, size=10)
        even_fill = PatternFill("solid", fgColor="F2F5F8")
        odd_fill = PatternFill("solid", fgColor="FFFFFF")
        border = Border(
            left=Side(style="thin", color="D3D3D3"),
            right=Side(style="thin", color="D3D3D3"),
            top=Side(style="thin", color="D3D3D3"),
            bottom=Side(style="thin", color="D3D3D3"),
        )

        # Header row
        ws.row_dimensions[1].height = 28
        for ci, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=ci)
            cell.fill = h_fill
            cell.font = h_font
            cell.border = border
            left_align = col_name in ("Scheme Name", "Asset Class")
            cell.alignment = Alignment(
                horizontal="left" if left_align else "center",
                vertical="center",
            )

        # Data rows
        for ri in range(2, len(df) + 2):
            ws.row_dimensions[ri].height = 20
            fill = even_fill if ri % 2 == 0 else odd_fill
            for ci, col_name in enumerate(df.columns, 1):
                cell = ws.cell(row=ri, column=ci)
                cell.fill = fill
                cell.font = data_font
                cell.border = border
                if col_name in date_cols:
                    if cell.value is not None:
                        cell.number_format = "0.0000"
                    else:
                        cell.value = "—"
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_name in ("Scheme Name", "Asset Class"):
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        # Column widths
        for col in ws.columns:
            max_len = max(
                (len(str(c.value or "")) + (3 if c.row == 1 else 0) for c in col),
                default=12,
            )
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(
                min(max_len + 2, 55), 12
            )

    return buf.getvalue()


@st.cache_data(ttl=900, show_spinner=False)
def fetch_amfi_data(frmdt: str, todt: str) -> pd.DataFrame:
    url = AMFI_HISTORY_URL.format(frmdt=frmdt, todt=todt)
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()
    return parse_amfi_text(resp.text)


# ─── UI ───────────────────────────────────────────────────────────────────────

def render_hero():
    st.markdown(
        """
        <div class="hero-card">
            <p class="hero-title">📊 AMFI NAV Fetcher</p>
            <p class="hero-sub">
                Fetch live &amp; historical NAV data from AMFI India — by ISIN or date range.
                Preview inline and export as a styled Excel report.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_date_cols(start_date, end_date, skip_sunday: bool) -> List[str]:
    """Generate sorted list of date strings between start and end."""
    cols = []
    current = start_date
    while current <= end_date:
        if not (skip_sunday and current.weekday() == 6):
            cols.append(current.strftime("%d-%b-%Y"))
        current += timedelta(days=1)
    return cols


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    render_hero()

    # ── Mode selector ─────────────────────────────────────────────────────────
    mode = st.radio(
        "Fetch mode",
        ["🔍 By ISIN + Date Range", "📅 By Date Range Only (all funds)"],
        horizontal=True,
    )

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Date inputs ───────────────────────────────────────────────────────────
    default_end = datetime.today().date()
    default_start = (datetime.today() - timedelta(days=7)).date()

    col_a, col_b = st.columns(2)
    with col_a:
        start_date = st.date_input("📅 Start Date", value=default_start)
    with col_b:
        end_date = st.date_input("📅 End Date", value=default_end)

    # ── ISIN input (only shown in ISIN mode) ──────────────────────────────────
    isin_list: List[str] = []
    if mode.startswith("🔍"):
        st.markdown(
            '<div class="mode-info">'
            "<strong>ISIN Mode:</strong> Enter one or more ISINs (Growth <em>or</em> Reinvestment) — "
            "one per line, or comma-separated. Rows matching either ISIN column will be fetched."
            "</div>",
            unsafe_allow_html=True,
        )
        raw_isins = st.text_area(
            "ISINs",
            height=160,
            placeholder="INF209K01AJ8\nINF846K01CH7\nINF760K01019\n...",
        )
        isin_list = [x.strip() for x in re.split(r"[,\n\s]+", raw_isins) if x.strip()]
    else:
        st.markdown(
            '<div class="mode-info">'
            "<strong>Date-Range Mode:</strong> No ISINs required — the full AMFI database "
            "for the selected period will be fetched (~8,000+ schemes). "
            "Large date ranges may take a few seconds."
            "</div>",
            unsafe_allow_html=True,
        )

    # ── Options ───────────────────────────────────────────────────────────────
    c1, c2, c3 = st.columns(3)
    with c1:
        skip_sunday = st.checkbox("Skip Sundays", value=True)
    with c2:
        carry_forward = st.checkbox("Carry forward on holidays", value=True)
    with c3:
        pivot_dates = st.checkbox("Pivot: one column per date", value=True)

    st.markdown("")

    fetch_btn = st.button("⚡ Fetch NAV Data", use_container_width=True)

    if not fetch_btn:
        return

    # ── Validation ────────────────────────────────────────────────────────────
    if start_date > end_date:
        st.error("Start Date must be before or equal to End Date.")
        return
    if mode.startswith("🔍") and not isin_list:
        st.error("Please enter at least one ISIN.")
        return

    # ── Fetch ─────────────────────────────────────────────────────────────────
    frmdt_str = start_date.strftime("%d-%b-%Y")
    todt_str = end_date.strftime("%d-%b-%Y")

    with st.spinner(f"Contacting AMFI India for {frmdt_str} → {todt_str} …"):
        try:
            df_raw = fetch_amfi_data(frmdt_str, todt_str)
        except Exception as exc:
            st.error(f"Failed to fetch data: {exc}")
            return

    if df_raw.empty:
        st.warning("No data returned from AMFI for this date range.")
        return

    # ── Filter by ISINs if needed ─────────────────────────────────────────────
    if mode.startswith("🔍"):
        df_filtered = filter_by_isins(df_raw, isin_list)
        if df_filtered.empty:
            st.warning(
                "No records found for the specified ISINs. "
                "Double-check the ISINs or widen the date range."
            )
            return
    else:
        df_filtered = df_raw

    # ── Build target date columns ─────────────────────────────────────────────
    date_cols = build_date_cols(start_date, end_date, skip_sunday)

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
        df_display = pivot_to_wide(df_filtered, date_cols)

        # Carry forward
        if carry_forward and len(date_cols) > 1:
            date_objs = sorted([datetime.strptime(d, "%d-%b-%Y") for d in date_cols])
            sorted_cols = [d.strftime("%d-%b-%Y") for d in date_objs]
            for i in range(1, len(sorted_cols)):
                prev, curr = sorted_cols[i - 1], sorted_cols[i]
                if prev in df_display.columns and curr in df_display.columns:
                    df_display[curr] = df_display[curr].fillna(df_display[prev])
    else:
        # Long format — show raw rows with selected columns
        wanted = [
            "Asset Class", "Scheme Code",
            "ISIN Div Payout / ISIN Growth", "ISIN Div Reinvestment",
            "Scheme Name", "Plan Type", "Option Type", "NAV", "NAV Date",
        ]
        df_display = df_filtered[[c for c in wanted if c in df_filtered.columns]].copy()
        date_cols = []  # no date pivot columns for long format

    # ── Preview ───────────────────────────────────────────────────────────────
    st.markdown("### 📋 Data Preview")
    st.dataframe(df_display, use_container_width=True, hide_index=True)

    # ── Export ────────────────────────────────────────────────────────────────
    with st.spinner("Generating styled Excel…"):
        excel_bytes = style_excel(df_display, date_cols)

    file_label = f"amfi_nav_{start_date}_to_{end_date}.xlsx"
    st.download_button(
        label="📥  Download Styled Excel (.xlsx)",
        data=excel_bytes,
        file_name=file_label,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # ── Summary table (long format only) ─────────────────────────────────────
    if not pivot_dates:
        with st.expander("📊 Scheme Summary", expanded=False):
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


if __name__ == "__main__":
    main()
