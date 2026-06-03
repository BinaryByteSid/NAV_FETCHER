import os
import pandas as pd

path = "c:/Users/sidha/OneDrive/Desktop/portfolio last 6 months.xlsx"
if os.path.exists(path):
    try:
        df = pd.read_excel(path, header=3)
        print("Unique PD_Month End values:")
        print(df['PD_Month End'].dropna().unique().tolist())
    except Exception as e:
        print(f"Error: {e}")
else:
    print("File not found")
