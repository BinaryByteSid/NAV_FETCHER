"""Find build_date_cols in nav_fetcher.py"""
with open("NAV_FETCHER/nav_fetcher.py", "r", encoding="utf-8", errors="ignore") as f:
    lines = f.read().splitlines()

for i, line in enumerate(lines, 1):
    if "def build_date_cols" in line:
        print(f"Found build_date_cols on line {i}: {line}")
        # Print next 20 lines
        for j in range(i, min(i+20, len(lines))):
            print(f"  Line {j+1}: {lines[j]}")
        break
