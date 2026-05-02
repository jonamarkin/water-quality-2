"""Read DECIMAL DATE sheet fully to find all column names."""
import pandas as pd, warnings
warnings.filterwarnings("ignore")

xl = pd.ExcelFile("parameters_used.xlsx", engine="openpyxl")

# DECIMAL DATE - read with first row as header
df = pd.read_excel(xl, sheet_name="DECIMAL DATE", header=0)
print("DECIMAL DATE columns:")
for i, c in enumerate(df.columns):
    print(f"  col {i}: {c!r}")
print("\nData:")
print(df.to_string())

# LEACHING RATE - read fully
print("\n\nLEACHING RATE - full column headers (row 1):")
df_lr = pd.read_excel(xl, sheet_name="LEACHING RATE ", header=None)
print("Row 1 (headers):", list(df_lr.iloc[1]))
print("Row 2 (2021 data):", list(df_lr.iloc[2]))
print("Row 25 (Lev mean g/ton):", list(df_lr.iloc[25]))

# PRODUCTION sheet
print("\n\nPRODUCTION:")
df_p = pd.read_excel(xl, sheet_name="PRODUCTION", header=0)
print(df_p.to_string())
