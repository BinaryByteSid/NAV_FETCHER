"""List all schemes in Large & Mid Cap (subcategory 2) on 31-Mar-2026"""
import requests

url = "https://www.amfiindia.com/gateway/pollingsebi/api/amfi/fundperformance"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Content-Type": "application/json"
}

payload = {
    "maturityType": 1,
    "category": 1,
    "subCategory": 2,
    "mfid": 0,
    "reportDate": "31-Mar-2026"
}
resp = requests.post(url, json=payload, headers=headers)
if resp.status_code == 200:
    data = resp.json().get("data", [])
    print(f"Total schemes in subCategory 2: {len(data)}")
    for r in data:
        name = r.get("schemeName")
        aum = r.get("dailyAUM")
        print(f"  Name: {name:<60} | AUM: {aum}")
else:
    print("Failed to fetch:", resp.status_code)
