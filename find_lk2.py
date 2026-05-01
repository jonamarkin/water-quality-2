"""
Explore the Process water sheet beyond row 200 to find LK blocks.
Also look at Flow-Production for LK production volumes.
"""
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

xl = pd.ExcelFile('Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx', engine='openpyxl')
df_pw = pd.read_excel(xl, sheet_name='Process water', header=None)
print(f"Process water sheet: {df_pw.shape[0]} rows x {df_pw.shape[1]} cols")

# Print ALL non-empty rows after row 192 (after the Ni block)
print("\n=== Process water rows 193 onwards (after Ni block) ===")
for i in range(193, df_pw.shape[0]):
    row = df_pw.iloc[i]
    non_null = [(j, row.iloc[j]) for j in range(df_pw.shape[1])
                if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() not in ('', 'nan')]
    if non_null:
        print(f"Row {i}: {non_null[:20]}")

# Flow-Production sheet
print("\n\n=== Flow-Production sheet ===")
df_fp = pd.read_excel(xl, sheet_name='Flow-Production', header=None)
print(f"Shape: {df_fp.shape}")
for i in range(df_fp.shape[0]):
    row = df_fp.iloc[i]
    non_null = [(j, row.iloc[j]) for j in range(df_fp.shape[1])
                if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() not in ('', 'nan')]
    if non_null:
        print(f"Row {i}: {non_null[:15]}")
