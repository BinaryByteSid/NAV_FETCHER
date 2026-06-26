"""Compare root app.py and NAV_FETCHER/app.py"""
import os

root_app = "app.py"
nf_app = "NAV_FETCHER/app.py"

if os.path.exists(root_app) and os.path.exists(nf_app):
    with open(root_app, "r", encoding="utf-8", errors="ignore") as f1:
        c1 = f1.read()
    with open(nf_app, "r", encoding="utf-8", errors="ignore") as f2:
        c2 = f2.read()
    if c1 == c2:
        print("Root app.py and NAV_FETCHER/app.py are IDENTICAL")
    else:
        print("Root app.py and NAV_FETCHER/app.py are DIFFERENT!")
        print(f"  Root app.py size: {len(c1)} characters")
        print(f"  NAV_FETCHER/app.py size: {len(c2)} characters")
else:
    print("One or both app.py files missing")
