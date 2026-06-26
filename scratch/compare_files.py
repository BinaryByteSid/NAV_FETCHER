"""Compare functions in NAV_FETCHER/app.py and NAV_FETCHER/nav_fetcher.py"""
import inspect
import sys, os
sys.path.insert(0, os.path.abspath("NAV_FETCHER"))
import app
import nav_fetcher

print("Comparing calculate_flows_for_dataframe:")
app_code = inspect.getsource(app.calculate_flows_for_dataframe)
nf_code = inspect.getsource(nav_fetcher.calculate_flows_for_dataframe)
if app_code == nf_code:
    print("  calculate_flows_for_dataframe is IDENTICAL")
else:
    print("  calculate_flows_for_dataframe is DIFFERENT!")
    
print("\nComparing populate_actual_aum:")
app_code_aum = inspect.getsource(app.populate_actual_aum)
nf_code_aum = inspect.getsource(nav_fetcher.populate_actual_aum)
if app_code_aum == nf_code_aum:
    print("  populate_actual_aum is IDENTICAL")
else:
    print("  populate_actual_aum is DIFFERENT!")
