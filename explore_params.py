"""Explore the parameters_used.xlsx file - all sheets and contents."""
import pandas as pd, warnings
warnings.filterwarnings("ignore")

xl = pd.ExcelFile("parameters_used.xlsx", engine="openpyxl")
print("Sheets:", xl.sheet_names)

for sh in xl.sheet_names:
    df = pd.read_excel(xl, sheet_name=sh, header=None)
    print(f"\n{'='*60}")
    print(f"Sheet: [{sh}]  Shape: {df.shape}")
    print(f"{'='*60}")
    # Print ALL rows that have any content
    for i in range(df.shape[0]):
        row = df.iloc[i]
        non_null = [(j, row.iloc[j]) for j in range(df.shape[1])
                    if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() not in ("", "nan")]
        if non_null:
            print(f"  Row {i}: {non_null}")
