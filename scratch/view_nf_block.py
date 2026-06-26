"""View nav_fetcher.py block around checkbox declarations"""
with open("NAV_FETCHER/nav_fetcher.py", "r", encoding="utf-8", errors="ignore") as f:
    lines = f.read().splitlines()

for idx in range(1060, min(1180, len(lines))):
    print(f"Line {idx+1}: {repr(lines[idx])}")
