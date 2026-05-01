"""Find and display all block headers in the full Process water sheet."""
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

xl = pd.ExcelFile('Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx', engine='openpyxl')
df = pd.read_excel(xl, sheet_name='Process water', header=None)
print(f"Total rows: {df.shape[0]}")

# Find all rows that contain 'Year' in col 0, or look like block headers
print("\n=== All header rows (contain 'Year' or block label) ===")
for i in range(df.shape[0]):
    row = df.iloc[i]
    v0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
    if v0 == 'Year':
        vals = [(j, row.iloc[j]) for j in range(min(12, df.shape[1]))
                if pd.notna(row.iloc[j])]
        print(f"Row {i}: {vals}")

# Also find label rows (2 rows above each 'Year' row typically have the parameter name)
print("\n=== Rows 193 to 220 (first block after Ni) ===")
for i in range(193, 225):
    row = df.iloc[i]
    non_null = [(j, row.iloc[j]) for j in range(min(15, df.shape[1]))
                if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() not in ('', 'nan')]
    if non_null:
        print(f"Row {i}: {non_null}")

print("\n=== Rows 350 to 380 ===")
for i in range(350, 385):
    row = df.iloc[i]
    non_null = [(j, row.iloc[j]) for j in range(min(12, df.shape[1]))
                if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() not in ('', 'nan')]
    if non_null:
        print(f"Row {i}: {non_null}")
