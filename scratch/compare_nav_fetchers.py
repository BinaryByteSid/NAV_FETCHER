"""Compare root nav_fetcher.py and NAV_FETCHER/nav_fetcher.py"""
import os

root_nf = "nav_fetcher.py"
nf_nf = "NAV_FETCHER/nav_fetcher.py"

if os.path.exists(root_nf) and os.path.exists(nf_nf):
    with open(root_nf, "r", encoding="utf-8", errors="ignore") as f1:
        c1 = f1.read()
    with open(nf_nf, "r", encoding="utf-8", errors="ignore") as f2:
        c2 = f2.read()
    if c1 == c2:
        print("Root nav_fetcher.py and NAV_FETCHER/nav_fetcher.py are IDENTICAL")
    else:
        print("Root nav_fetcher.py and NAV_FETCHER/nav_fetcher.py are DIFFERENT!")
        print(f"  Root nav_fetcher.py size: {len(c1)} characters")
        print(f"  NAV_FETCHER/nav_fetcher.py size: {len(c2)} characters")
else:
    print("One or both nav_fetcher.py files missing")
