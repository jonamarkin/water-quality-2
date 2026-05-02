"""Verify row ranges and storage column availability for all 12 blocks."""
import pandas as pd, numpy as np, warnings
warnings.filterwarnings("ignore")

df = pd.read_excel("Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx",
                   sheet_name="Process water", header=None, engine="openpyxl")
print(f"Total rows: {df.shape[0]}, cols: {df.shape[1]}")

CANDIDATE_BLOCKS = {
    "Cu":   (42,  74),
    "NH4":  (81,  113),
    "Cl":   (121, 153),
    "Ni":   (160, 192),
    "Zn":   (199, 231),
    "Co":   (239, 271),
    "Mo":   (278, 310),
    "SO4":  (320, 352),
    "Ca":   (360, 392),
    "NO3":  (396, 428),
    "As":   (506, 538),
    "Cr":   (543, 575),
}

print(f"\n{'Param':<8} {'First':>6} {'Last':>6} {'Data rows':>10} "
      f"{'Col9 (mod)':>12} {'Col31 (store)':>14} {'Max col':>8}")
print("-"*70)

for param, (fr, lr) in CANDIDATE_BLOCKS.items():
    rows = df.iloc[fr:lr+1]
    year_rows = rows[rows.iloc[:, 0].apply(
        lambda v: isinstance(v, (int,float)) and 2013 <= float(v) <= 2032
    )]
    n = len(year_rows)

    # Check col 9 (modelled) and col 31 (storage)
    c9_ok  = year_rows.iloc[:,9].notna().sum() if df.shape[1] > 9 else 0
    c31_ok = year_rows.iloc[:,31].notna().sum() if df.shape[1] > 31 else 0
    max_col = year_rows.shape[1]

    # Sample first data row values
    if len(year_rows) > 0:
        r0 = year_rows.iloc[0]
        c9v  = f"{r0.iloc[9]:.4f}"  if df.shape[1]>9  and pd.notna(r0.iloc[9])  else "NaN"
        c31v = f"{r0.iloc[31]:.4f}" if df.shape[1]>31 and pd.notna(r0.iloc[31]) else "NaN"
    else:
        c9v = c31v = "?"

    print(f"{param:<8} {fr:>6} {lr:>6} {n:>10} "
          f"{c9v:>12} {c31v:>14} {max_col:>8}")
