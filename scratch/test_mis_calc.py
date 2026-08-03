"""
Scratch unit test script for MIS Generator calculations and exports.

Hits the live AMFI feed, so it takes a few minutes. Assertions cover structure
and internal consistency rather than specific NAV values.
"""

from __future__ import annotations
import sys
from datetime import date
from pathlib import Path

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd
from mis_generator import (
    validate_and_normalize_portfolio,
    generate_mis_reports_data,
    export_mis_to_excel,
    export_mis_to_pdf,
    DEFAULT_SAMPLE_PORTFOLIO,
    MAX_DAILY_GAP_DAYS,
)
from benchmark_proxy import normalize_benchmark_name, STATUS_NONE


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
    df = test_validation()
    res = test_reports(df)
    test_exports(res)
    if res["warnings"]:
        print("\nWarnings raised by this run:")
        for w in res["warnings"]:
            print("  *", w)
    print("\nALL MIS GENERATOR TESTS PASSED.")
