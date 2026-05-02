"""Print ALL sheet names and structure of parameters_used.xlsx."""
import pandas as pd, warnings
warnings.filterwarnings("ignore")

xl = pd.ExcelFile("parameters_used.xlsx", engine="openpyxl")
print("ALL SHEETS:", xl.sheet_names)
print()

for sh in xl.sheet_names:
    df = pd.read_excel(xl, sheet_name=sh, header=None)
    print(f"=== [{sh}]  {df.shape[0]} rows x {df.shape[1]} cols ===")
    # Print first 5 non-empty rows to get structure
    count = 0
    for i in range(df.shape[0]):
        row = df.iloc[i]
        non_null = [(j, row.iloc[j]) for j in range(df.shape[1])
                    if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() not in ("","nan")]
        if non_null:
            print(f"  Row {i}: {non_null[:12]}")
            count += 1
        if count >= 6:
            print("  ...")
            break
    print()
