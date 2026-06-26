"""Quick verification that parse_amfi_date_series works correctly."""
import sys, os
sys.path.insert(0, os.path.abspath("NAV_FETCHER"))
import pandas as pd

# Simulate the helper
_MONTH_ABBR = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04",
    "may": "05", "jun": "06", "jul": "07", "aug": "08",
    "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}

def _parse_amfi_date_str(date_str):
    if not isinstance(date_str, str) or not date_str.strip():
        return None
    parts = date_str.strip().split("-")
    if len(parts) == 3:
        day, month_abbr, year = parts
        month_num = _MONTH_ABBR.get(month_abbr.lower()[:3])
        if month_num and day.isdigit() and year.isdigit():
            return f"{year}-{month_num}-{day.zfill(2)}"
    return None

# Test with AMFI format dates
test_dates = pd.Series([
    "01-Jun-2026", "02-Jun-2026", "03-Jun-2026",
    "29-May-2026", "30-May-2026", "31-May-2026",
    "15-Jan-2026", "28-Feb-2026", "10-Dec-2025"
])

print("Test dates:")
print(test_dates.tolist())
print()

# Test pd.to_datetime with format
result1 = pd.to_datetime(test_dates, format="%d-%b-%Y", errors="coerce")
print("pd.to_datetime(format='%d-%b-%Y'):")
print(result1.tolist())
print(f"  NaT count: {result1.isna().sum()}")
print()

# Test manual parsing
iso_strs = test_dates.apply(_parse_amfi_date_str)
result2 = pd.to_datetime(iso_strs, errors="coerce")
print("Manual parse -> ISO -> pd.to_datetime:")
print(result2.tolist())
print(f"  NaT count: {result2.isna().sum()}")
print()

# Test with pre-converted dd-mm-YYYY format
converted = pd.Series(["01-06-2026", "02-06-2026", "03-06-2026"])
result3 = pd.to_datetime(converted, format="%d-%m-%Y", errors="coerce")
print("pd.to_datetime(format='%d-%m-%Y') for dd-mm-YYYY:")
print(result3.tolist())
print(f"  NaT count: {result3.isna().sum()}")

# Now test the actual function from nav_fetcher
from nav_fetcher import parse_amfi_date_series
result4 = parse_amfi_date_series(test_dates)
print("\nparse_amfi_date_series():")
print(result4.tolist())
print(f"  NaT count: {result4.isna().sum()}")
print()

# Verify all dates are unique and different
strs = result4.dt.strftime("%d-%m-%Y")
print("Formatted as dd-mm-YYYY:")
print(strs.tolist())
print(f"  Unique: {strs.nunique()}")
print("\n✅ All dates parsed correctly!" if result4.isna().sum() == 0 else "\n❌ Some dates failed!")
