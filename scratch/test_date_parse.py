import pandas as pd
import sys

print("Pandas version:", pd.__version__)
print("Python version:", sys.version)

dates = ["29-May-2026", "31-May-2026", "01-Jun-2026", "02-Jun-2026"]
for d in dates:
    res = pd.to_datetime(d, errors="coerce")
    print(f"Simple pd.to_datetime('{d}') -> {res}")
    
    res_fmt = pd.to_datetime(d, format="%d-%b-%Y", errors="coerce")
    print(f"Format pd.to_datetime('{d}', format='%d-%b-%Y') -> {res_fmt}")
