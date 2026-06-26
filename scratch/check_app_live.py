"""Check fetch_live_aum in app.py"""
with open("NAV_FETCHER/app.py", "r", encoding="utf-8", errors="ignore") as f:
    code = f.read()

for i, line in enumerate(code.splitlines(), 1):
    if "fetch_live_aum" in line or "live daily AUM" in line:
        print(f"Line {i}: {line}")
