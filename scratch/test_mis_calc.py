"""
Scratch unit test script for MIS Generator calculations and exports.

Hits the live AMFI feed, so it takes a few minutes. Assertions cover structure
and internal consistency rather than specific NAV values.
"""

from __future__ import annotations
import sys
import tempfile
from datetime import date
from pathlib import Path

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).parent.parent))

import openpyxl
import pandas as pd
from mis_generator import (
    validate_and_normalize_portfolio,
    read_portfolio_excel,
    _is_percent_format,
    generate_mis_reports_data,
    export_mis_to_excel,
    export_mis_to_pdf,
    DEFAULT_SAMPLE_PORTFOLIO,
    MAX_DAILY_GAP_DAYS,
    NIFTY_KEY,
    _build_reports,
)
from benchmark_proxy import normalize_benchmark_name, BenchmarkSeries, STATUS_NONE, STATUS_EXACT


def test_benchmark_normalization():
    print("1. Benchmark name normalization...")
    cases = {
        "S&P BSE 250 SmallCap TR INR": "BSE 250 SMALLCAP",
        "BSE 500 India TR INR": "BSE 500",
        "Nifty Midcap 150 TR INR": "NIFTY MIDCAP 150",
        "Nifty LargeMidcap 250 TR INR": "NIFTY LARGEMIDCAP 250",
        "Nifty 500 Multicap 50:25:25 TR INR": "NIFTY 500 MULTICAP 50:25:25",
        "NIFTY 50": "NIFTY 50",
    }
    for raw, want in cases.items():
        got = normalize_benchmark_name(raw)
        assert got == want, f"{raw!r} -> {got!r}, expected {want!r}"
    print("   [OK] All benchmark labels normalized correctly.")


def test_validation():
    print("2. Portfolio validation & normalization...")
    clean_df, warnings, errors = validate_and_normalize_portfolio(pd.DataFrame(DEFAULT_SAMPLE_PORTFOLIO))
    assert not errors, f"Unexpected errors: {errors}"
    assert len(clean_df) == 14, f"Expected 14 rows, got {len(clean_df)}"
    assert abs(clean_df["Normalized_Weight"].sum() - 1.0) < 1e-9
    print("   [OK] Validation passed.")
    return clean_df


def test_excel_percent_allocations():
    """A cell holding 7 always reads as 7. Excel's percent format is the only
    thing that turns a stored 0.07 into 7%."""
    print("2b. Excel percent-formatted allocations...")

    assert _is_percent_format("0.00%")
    assert _is_percent_format("0%")
    assert not _is_percent_format("General")
    assert not _is_percent_format("0.00")
    # A '%' that is literal text, not a percent multiplier.
    assert not _is_percent_format('0.00" %"')
    assert not _is_percent_format(r"0.00\%")

    def write_book(path, allocs, number_format):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Scheme Name", "ISIN", "Allocation", "Benchmark"])
        for i, a in enumerate(allocs):
            ws.append([f"F{i}", f"INF{i:03d}X01XX{i % 10}", a, "NIFTY 50"])
            ws.cell(row=i + 2, column=3).number_format = number_format
        wb.save(path)
        return path

    tmp = Path(tempfile.gettempdir())

    # Percent-formatted: 0.07 displays as 7.00%, so it reads as 7.
    p = write_book(tmp / "_mis_pct.xlsx", [0.07, 0.10, 0.09, 0.74], "0.00%")
    with open(p, "rb") as fh:
        df = read_portfolio_excel(fh)
    assert list(df["Allocation"]) == [7.0, 10.0, 9.0, 74.0], list(df["Allocation"])

    # Plain format: 7 is 7, and stays 7.
    p = write_book(tmp / "_mis_plain.xlsx", [50.0, 30.0, 20.0], "General")
    with open(p, "rb") as fh:
        df = read_portfolio_excel(fh)
    assert list(df["Allocation"]) == [50.0, 30.0, 20.0], list(df["Allocation"])

    # 100 holdings at 1% each, stored plainly: untouched.
    p = write_book(tmp / "_mis_ones.xlsx", [1.0] * 100, "General")
    with open(p, "rb") as fh:
        df = read_portfolio_excel(fh)
    assert df["Allocation"].iloc[0] == 1.0, df["Allocation"].iloc[0]
    assert df["Allocation"].sum() == 100.0

    # The validator itself never rescales: what it is handed is what it uses.
    lit = pd.DataFrame({
        "Scheme Name": ["A", "B"],
        "ISIN": ["INF001A01AA1", "INF002B01BB2"],
        "Allocation (%)": [0.07, 0.10],
        "Benchmark": ["NIFTY 50"] * 2,
    })
    clean, _, _ = validate_and_normalize_portfolio(lit)
    assert list(clean["Allocation (%)"]) == [0.07, 0.10], list(clean["Allocation (%)"])

    # Text cells carrying their own percent sign.
    txt = pd.DataFrame({
        "Scheme Name": ["A", "B"],
        "ISIN": ["INF001A01AA1", "INF002B01BB2"],
        "Allocation (%)": ["60%", "40%"],
        "Benchmark": ["NIFTY 50"] * 2,
    })
    c4, _, _ = validate_and_normalize_portfolio(txt)
    assert list(c4["Allocation (%)"]) == [60.0, 40.0], list(c4["Allocation (%)"])
    print("   [OK] Percent-formatted allocations handled.")


def test_period_baseline_is_prior_close():
    """A period return compounds its own first day, so it bases off the close
    BEFORE the period opens. Basing off the start date itself drops day one and
    understates every return."""
    print("2c. Period/FY baseline convention...")

    # 31 Mar = 100, then +10% on 1 Apr, then flat to the report date.
    series = {
        date(2026, 3, 31): 100.0,
        date(2026, 4, 1): 110.0,
        date(2026, 7, 22): 110.0,
        date(2026, 7, 23): 110.0,
    }
    nav = {"INF000A01AA1": series}
    bm = {NIFTY_KEY: BenchmarkSeries(NIFTY_KEY, NIFTY_KEY, series, STATUS_EXACT)}
    port = pd.DataFrame([{
        "Scheme Name": "F", "ISIN": "INF000A01AA1", "Allocation (%)": 100.0,
        "Benchmark": NIFTY_KEY, "Normalized_Weight": 1.0,
    }])
    flows = {"INF000A01AA1": {"day": None, "mtd": None, "ytd": None}}

    block = _build_reports(
        port, nav, bm, flows,
        d_end=date(2026, 7, 23), d_start=date(2026, 4, 1),
        d_fy=date(2026, 4, 1), d_mtd=date(2026, 6, 30), label="T",
    )
    row = block["mis3"]["rows"].iloc[0]
    # From 31 Mar's close of 100 to 110 is 10%. Basing off 1 Apr gives 0%.
    assert abs(row["FY Scheme Return"] - 10.0) < 1e-6, row["FY Scheme Return"]

    row1 = block["mis1"]["rows"].iloc[0]
    assert abs(row1["Period Scheme Return"] - 10.0) < 1e-6, row1["Period Scheme Return"]
    print("   [OK] Baselines take the prior close.")


def test_reports(clean_df):
    print("3. MIS report generation (live AMFI fetch, please wait)...")
    results = generate_mis_reports_data(clean_df, date(2026, 4, 1), date(2026, 7, 31))

    for block_key in ("mis1", "mis2", "mis3"):
        spec = results["current"][block_key]
        assert not spec["rows"].empty, f"{block_key} produced no rows"
        assert len(spec["rows"]) == 14, f"{block_key} lost rows"
        assert set(spec["columns"]) <= set(spec["headers"]), f"{block_key} header/column mismatch"
    print("   [OK] All three MIS tables built.")

    # Excess return must equal scheme minus comparator wherever both exist.
    detail = results["current"]["detail"]
    triples = [
        ("Day Scheme Return", "Day Nifty Return", "Day Excess vs Nifty"),
        ("Period Scheme Return", "Period Nifty Return", "Period Excess vs Nifty"),
        ("Period Scheme Return", "Period Benchmark Return", "Period Excess vs Benchmark"),
        ("FY Scheme Return", "FY Nifty Return", "FY Excess vs Nifty"),
        ("FY Scheme Return", "FY Benchmark Return", "FY Excess vs Benchmark"),
    ]
    checked = 0
    for _, r in detail.iterrows():
        for a, b, x in triples:
            if pd.notna(r[a]) and pd.notna(r[b]) and pd.notna(r[x]):
                assert abs((r[a] - r[b]) - r[x]) < 1e-9, f"excess mismatch on {r['Scheme Name']} / {x}"
                checked += 1
    assert checked > 0, "No excess returns were verifiable — check the data feed"
    print(f"   [OK] Excess-return identity holds across {checked} cells.")

    # A one-day return must come from consecutive observations, never a
    # multi-week gap dressed up as a day.
    for _, r in detail.iterrows():
        if pd.notna(r["Day Scheme Return"]):
            gap = r["_day_gap_days"]
            assert gap is not None and gap <= MAX_DAILY_GAP_DAYS, \
                f"{r['Scheme Name']} reports a daily return across a {gap}-day gap"
    print("   [OK] Daily returns all span consecutive observations.")

    # Benchmarks must never be silently invented.
    for _, r in results["benchmark_report"].iterrows():
        if r["Status"] == STATUS_NONE:
            assert r["Observations"] == 0, "Unavailable benchmark reported observations"
        else:
            assert r["Proxy ISIN"] != "-", f"{r['Benchmark']} resolved without a proxy fund"
    print("   [OK] Every benchmark traces to a named proxy fund or reports N/A.")

    return results


def test_exports(results):
    print("4. Excel export...")
    xl = export_mis_to_excel(results)
    assert len(xl) > 0
    print(f"   [OK] Excel generated ({len(xl)} bytes)")

    print("5. PDF export...")
    pdf = export_mis_to_pdf(results)
    assert len(pdf) > 0
    print(f"   [OK] PDF generated ({len(pdf)} bytes)")


if __name__ == "__main__":
    test_benchmark_normalization()
    test_excel_percent_allocations()
    test_period_baseline_is_prior_close()
    df = test_validation()
    res = test_reports(df)
    test_exports(res)
    if res["warnings"]:
        print("\nWarnings raised by this run:")
        for w in res["warnings"]:
            print("  *", w)
    print("\nALL MIS GENERATOR TESTS PASSED.")
