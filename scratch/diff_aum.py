"""Diff populate_actual_aum between app.py and nav_fetcher.py"""
import inspect
import sys, os
import difflib
sys.path.insert(0, os.path.abspath("NAV_FETCHER"))
import app
import nav_fetcher

app_code = inspect.getsource(app.populate_actual_aum).splitlines()
nf_code = inspect.getsource(nav_fetcher.populate_actual_aum).splitlines()

diff = difflib.unified_diff(app_code, nf_code, fromfile="app.py", tofile="nav_fetcher.py")
print("\n".join(diff))
