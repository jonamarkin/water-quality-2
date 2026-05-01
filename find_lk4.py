"""Map ALL block headers in the Process water sheet."""
import pandas as pd, warnings
warnings.filterwarnings('ignore')

df = pd.read_excel('Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx',
                   sheet_name='Process water', header=None, engine='openpyxl')
print(f"Total rows: {df.shape[0]}")

print("\n=== All rows with 'Year' in col 0 or 1 (block headers) ===")
for i in range(df.shape[0]):
    v0 = str(df.iloc[i, 0]).strip()
    v1 = str(df.iloc[i, 1]).strip() if df.shape[1] > 1 else ''
    if v0 == 'Year' or v1 == 'Tailings':
        # Look 2 rows above for the parameter name
        lbl = ''
        for back in range(1, 5):
            if i - back >= 0:
                row_above = df.iloc[i - back]
                for j in range(df.shape[1]):
                    v = str(row_above.iloc[j]).strip()
                    if v not in ('nan', '', 'True', 'False') and not v.replace('.','').isdigit():
                        lbl = v
                        break
                if lbl:
                    break
        print(f"  Header row {i}  -->  parameter hint: '{lbl}'")
        # Also show col 9 label (the modelled concentration unit label)
        c9 = str(df.iloc[i, 9]).strip() if df.shape[1] > 9 else '?'
        print(f"    col9={c9}")
