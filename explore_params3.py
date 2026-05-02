"""Extract the key data we need: DECIMAL DATE sheet and LEACHING RATE sheet."""
import pandas as pd, numpy as np, warnings
warnings.filterwarnings("ignore")

xl = pd.ExcelFile("parameters_used.xlsx", engine="openpyxl")
print("ALL SHEETS:", xl.sheet_names)

# ── 1. DECIMAL DATE sheet: actual measured concentrations 2016-2025 by half-year ──
print("\n\n=== DECIMAL DATE sheet (actual monitoring data) ===")
df_dd = pd.read_excel(xl, sheet_name="DECIMAL DATE", header=0)
print(df_dd.to_string())

# ── 2. LEACHING RATE sheet ──
print("\n\n=== LEACHING RATE sheet ===")
df_lr = pd.read_excel(xl, sheet_name="LEACHING RATE ", header=None)
for i in range(df_lr.shape[0]):
    row = df_lr.iloc[i]
    non_null = [(j, row.iloc[j]) for j in range(df_lr.shape[1])
                if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() not in ("","nan")]
    if non_null:
        print(f"  Row {i}: {non_null}")

# ── 3. PRODUCTION sheet ──
print("\n\n=== PRODUCTION sheet ===")
df_prod = pd.read_excel(xl, sheet_name="PRODUCTION", header=None)
for i in range(df_prod.shape[0]):
    row = df_prod.iloc[i]
    non_null = [(j, row.iloc[j]) for j in range(df_prod.shape[1])
                if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() not in ("","nan")]
    if non_null:
        print(f"  Row {i}: {non_null}")

# ── 4. Statistics/mix sheet ──
print("\n\n=== First sheet (mix data) ===")
sh0 = xl.sheet_names[0]
df0 = pd.read_excel(xl, sheet_name=sh0, header=None)
for i in range(df0.shape[0]):
    row = df0.iloc[i]
    non_null = [(j, row.iloc[j]) for j in range(df0.shape[1])
                if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() not in ("","nan")]
    if non_null:
        print(f"  Row {i}: {non_null}")
