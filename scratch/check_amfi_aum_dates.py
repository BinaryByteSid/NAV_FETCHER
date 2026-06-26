"""Check dailyAUM returned by AMFI for multiple dates"""
import requests

url = "https://www.amfiindia.com/gateway/pollingsebi/api/amfi/fundperformance"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Content-Type": "application/json"
}

dates = ["31-Mar-2026", "01-Apr-2026", "02-Apr-2026"]
# subCategory = 4 is Multi Cap Fund
# maturityType = 1, category = 1
for date in dates:
    payload = {
        "maturityType": 1,
        "category": 1,
        "subCategory": 4,
        "mfid": 0,
        "reportDate": date
    }
    resp = requests.post(url, json=payload, headers=headers)
    if resp.status_code == 200:
        data = resp.json().get("data", [])
        # Find Axis Multicap Fund
        axis_rows = [r for r in data if "Axis Multicap" in r.get("schemeName", "")]
        print(f"Date: {date}")
        for r in axis_rows:
            print(f"  Scheme: {r.get('schemeName')} | AUM: {r.get('dailyAUM')} | reportDate: {r.get('reportDate')}")
    else:
        print(f"Failed to fetch for {date}: {resp.status_code}")
