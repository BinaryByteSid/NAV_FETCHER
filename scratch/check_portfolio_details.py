"""Search for flexi cap funds in portfolio last 6 months.xlsx"""
import pandas as pd
import os

path = "c:/Users/sidha/OneDrive/Desktop/portfolio last 6 months.xlsx"
if os.path.exists(path):
    df = pd.read_excel(path, header=3)
    # Print columns
    print("Columns:", list(df.columns))
    # Filter schemes containing 'Flexi' or 'Focused' or any of our targets
    flexi_rows = df[df['Scheme Name'].astype(str).str.contains('Flexi|Focused|360 ONE|Abakkus|Aditya Birla|Canara|Axis|Bandhan', case=False, na=False)]
    print("\nMatching rows:")
    for _, r in flexi_rows.iterrows():
        print(f"Name: {r['Scheme Name']} | ISIN: {r['SD_Scheme ISIN']} | AUM: {r['PD_Scheme AUM']}")
else:
    print("Portfolio file not found at", path)
