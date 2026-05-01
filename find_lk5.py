"""Print 2 rows above each header to get the parameter name."""
import pandas as pd, warnings
warnings.filterwarnings('ignore')

df = pd.read_excel('Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx',
                   sheet_name='Process water', header=None, engine='openpyxl')

header_rows = [19,41,80,120,159,198,238,277,319,359,395,431,468,505,542]

for hr in header_rows:
    print(f"\n--- Block at header row {hr} ---")
    for back in range(3, 0, -1):
        r = hr - back
        row = df.iloc[r]
        non_null = [(j, row.iloc[j]) for j in range(min(12,df.shape[1]))
                    if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() not in ('','nan')]
        if non_null:
            print(f"  row {r}: {non_null}")
    # Show col9 header
    c9 = df.iloc[hr, 9]
    c4 = df.iloc[hr, 4] if df.shape[1]>4 else ''
    c6 = df.iloc[hr, 6] if df.shape[1]>6 else ''
    c7 = df.iloc[hr, 7] if df.shape[1]>7 else ''
    print(f"  [col9={c9}] [col4={c4}] [col6={c6}] [col7={c7}]")
    # First data row
    dr = hr + 1
    row = df.iloc[dr]
    print(f"  first data row {dr}: year={row.iloc[0]}, half={row.iloc[1]}, col4={row.iloc[4]:.2f}, col6={row.iloc[6]:.2f}, col7={row.iloc[7]:.2f}, col9={row.iloc[9]:.4f}")
