"""Check AMFI response and matching for Axis Multicap on 30-Mar-2026"""
import requests
import sys, os
sys.path.insert(0, os.path.abspath("NAV_FETCHER"))
import app

url = "https://www.amfiindia.com/gateway/pollingsebi/api/amfi/fundperformance"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Content-Type": "application/json"
}

payload = {
    "maturityType": 1,
    "category": 1,
    "subCategory": 4, # Multi Cap Fund
    "mfid": 0,
    "reportDate": "30-Mar-2026"
}
resp = requests.post(url, json=payload, headers=headers)
if resp.status_code == 200:
    data = resp.json().get("data", [])
    axis_rows = [r for r in data if "Axis Multicap" in r.get("schemeName", "")]
    print("For 30-Mar-2026:")
    if not axis_rows:
        print("  No Axis Multicap schemes found in AMFI response!")
        print("  Available schemes in subCategory 4:")
        for r in data[:10]:
            print(f"    - {r.get('schemeName')}")
    else:
        for r in axis_rows:
            match = app.find_matching_perf_row("Axis Multicap Fund - Regular Plan - Growth", [r])
            print(f"  Scheme: {r.get('schemeName')} | AUM: {r.get('dailyAUM')} | Matched?: {match is not None}")
else:
    print("HTTP error:", resp.status_code)
