"""Search all Excel sheets for LK / Kiruna / leaching rate data."""
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

xl = pd.ExcelFile('Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx', engine='openpyxl')
print('Sheets:', xl.sheet_names)

KEYWORDS = ['kiruna', 'lk', 'leveaniemi', 'leach', 'lev-kir', 'lev.kir']

for sh in xl.sheet_names:
    df = pd.read_excel(xl, sheet_name=sh, header=None)
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            v = df.iloc[i, j]
            if pd.notna(v):
                sv = str(v).strip().lower()
                if any(k in sv for k in KEYWORDS):
                    print(f"  Sheet [{sh}] row {i} col {j}: '{v}'")

print("\n--- ore-tail leach calc sheet (all rows) ---")
df2 = pd.read_excel(xl, sheet_name='ore-tail leach calc', header=None)
print(f"Shape: {df2.shape}")
for i in range(df2.shape[0]):
    row = df2.iloc[i]
    non_null = [(j, row.iloc[j]) for j in range(df2.shape[1]) if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() != '']
    if non_null:
        print(f"  Row {i}: {non_null}")
