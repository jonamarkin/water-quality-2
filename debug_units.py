"""Investigate scale of NH4 and Cl actual vs predicted."""
import pandas as pd, numpy as np, warnings
warnings.filterwarnings("ignore")

xl = pd.ExcelFile("parameters_used.xlsx", engine="openpyxl")
df = pd.read_excel(xl, sheet_name="DECIMAL DATE", header=0)

print("=== ACTUAL measured values ===")
print(df[["Period","Season","Cu","NH4-N","Cl","Ni"]].to_string())

# What is the NH4 scale in the consultant's model?
# From ore-tail leach calc: NH4 Kiruna = 3047.3 kg/Mton
# Production ~3 Mton/half-year summer = 3047.3 * 3 = 9141.9 kg NH4 added
# Process water volume ~25.21 Mm3 = 25,210,000 m3 = 25,210,000,000 L
# Concentration = 9141.9 kg / 25,210,000,000 L = 0.000000363 kg/L = 0.000363 mg/L
# That's incredibly tiny - the consultant's NH4 prediction was ~4 mg/L
# Something is off with units somewhere

# Let's check what the consultant's model actually predicts for NH4
# From our earlier run: NH4 P50 range 0.025 -- 1.613 mg/L (correct scale)
# The actual data: NH4-N is 0.009 to 0.179 mg/L (matching scale)

# The issue is the leaching rate. The consultant used 598 kg/Mton for NH4 GM
# and 3047 for GK. These give:
# 3047 kg/Mton * 3.2 Mton = 9750 kg NH4
# In 25.21 Mm3 = 25.21e9 L → 9750 kg / 25.21e9 L = 3.87e-7 kg/L = 0.000387 mg/L
# But the consultant's model shows 4 mg/L for NH4...
# So there's a missing factor of ~10000x somewhere

# Let's look at the formula: mass_in / vol_denom
# vol_denom in Mm3, leach in kg
# C = kg / Mm3 = kg / (1e6 m3) = kg / (1e9 L) = 1e-9 kg/L = 1e-6 mg/L
# That's way too small. The consultant must be working in different units
# Let's check the consultant's column units:
# COL_LEACH_GM = 4: "Process leaching GM" - from our output: values like 49.0, 174.5, 1744.17
# These are not in kg - they must be in different units

# From first explore: Row 81 (NH4 first data): col4=1744.17 (leach_gm)
# and col5=2.917 (prod_gm Mton), col6=3047 (leach_gk rate in col header row)
# Wait, col4 = 1744.17 seems to be in units where it's directly used in formula
# with pit pump in Mm3 and concentrations in mg/L
# So 1744.17 / vol_denom (Mm3) should give mg/L
# If vol_denom ~3 Mm3: 1744.17/3 = 581 mg/L... still too high

# Let's look at actual formula: from the consultant col4 header for NH4 block
# From find_lk5.py: "first data row 81: col4=1744.17"
# with production 2.917 Mton at 598 kg/Mton = 1744 kg <- matches!
# And in formula: 1744 kg / vol_denom Mm3 = 1744 kg / (x Mm3)
# 1 Mm3 = 1e6 m3 = 1e9 L
# So 1744 kg / (3 * 1e9 L) = 5.8e-7 kg/L = 0.00058 mg/L -- too small!
# But consultant shows 4.13 mg/L...

# CONCLUSION: The Q values (pit pump, volumes) must also be in units that cancel properly
# Let's check pit pump: COL_PIT_PUMP = 8: values like 1.017, 1.119 Mm3
# And storage vol AE: col 30 (AE)? Let's check what AE actually is

print("\n\n=== Checking consultant's NH4 formula dimensions ===")
xl2 = pd.ExcelFile("Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx", engine="openpyxl")
df_pw = pd.read_excel(xl2, sheet_name="Process water", header=None)

# NH4 block row 80 (0-based) - first data row is 81
# Check ALL non-null cols in rows 79-85
print("NH4 block rows 79-85:")
for i in range(79, 87):
    row = df_pw.iloc[i]
    non_null = [(j, round(float(row.iloc[j]),4) if isinstance(row.iloc[j], (int,float)) else row.iloc[j])
                for j in range(min(35, df_pw.shape[1]))
                if pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() not in ("","nan")]
    if non_null:
        print(f"  Row {i}: {non_null}")
