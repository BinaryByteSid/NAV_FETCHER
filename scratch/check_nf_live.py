"""Check fetch_live_aum occurrences in nav_fetcher.py"""
import re

with open("NAV_FETCHER/nav_fetcher.py", "r", encoding="utf-8", errors="ignore") as f:
    code = f.read()

for i, line in enumerate(code.splitlines(), 1):
    if "fetch_live_aum" in line:
        print(f"Line {i}: {line}")
