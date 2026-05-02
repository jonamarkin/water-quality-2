import warnings; warnings.filterwarnings('ignore')
import pandas as pd, numpy as np

SEP_UNITS = {
    'SO4': 'mg/L', 'Ca': 'mg/L', 'Cl': 'mg/L',
    'NO3': 'mg/L', 'NH4': 'mg/L',
    'Cu':  'ug/L', 'Ni': 'ug/L', 'Zn': 'ug/L',
    'Co':  'ug/L', 'Mo': 'ug/L', 'As': 'ug/L', 'Cr': 'ug/L',
}
SEP_2020 = {
    'SO4': 122.0, 'Ca': 79.4, 'Cl': 28.4,
    'NO3': 1.747, 'NH4': 0.5348,
    'Cu': 10.81, 'Ni': 0.8817, 'Zn': 3.806,
    'Co': 0.3771, 'Mo': 6.384, 'As': 0.1199, 'Cr': 0.0398,
}
# DECIMAL DATE monitoring 2020 annual avg
MON_2020 = {
    'SO4': 1071.0, 'Ca': 420.8, 'Cl': 168.5,
    'NO3': 8.73,   'NH4': 0.028,
    'Cu':  1.805,  'Ni': 1.427, 'Zn': 2.25,
    'Co':  0.466,  'Mo': 16.75, 'As': 0.270, 'Cr': 0.078,
}
UNITS_MON = {
    'SO4': 'mg/L', 'Ca': 'mg/L', 'Cl': 'mg/L',
    'NO3': 'mg/L', 'NH4': 'mg/L',
    'Cu': 'ug/L', 'Ni': 'ug/L', 'Zn': 'ug/L',
    'Co': 'ug/L', 'Mo': 'ug/L', 'As': 'ug/L', 'Cr': 'ug/L',
}

Q_pit   = 1.649414 / 2
Q_gb    = 0.975697 / 2
Q_disch = 3.001672 / 2
Q_leak  = 0.8 / 2
Q_out   = Q_disch + Q_leak

LEV = {'SO4':260.0,'Ca':141.8,'Cl':154.4,'NO3':0.0,'NH4':0.0,
       'Cu':0.001358,'Ni':0.000631,'Zn':0.005307,'Co':0.000107,'Mo':0.022956,'As':0.012511,'Cr':0.000163}
KIR = {'SO4':1653.9,'Ca':581.4,'Cl':285.4,'NO3':17.3,'NH4':3.047,
       'Cu':0.011668,'Ni':0.001174,'Zn':0.150454,'Co':0.000283,'Mo':0.133563,'As':0.012995,'Cr':0.000244}

prod = 4.2e6 * 7 / 12
frac_lev, frac_kir = 0.50, 0.42

print(f"Q_out = {Q_out:.3f} Mm3/half-yr  ({Q_out*1e9:.3e} L)")
print()
hdr = f"{'Param':4s}  {'LeachLoad_g':>12s}  {'PitLoad_g':>12s}  {'GBload_g':>12s}  {'Predicted':>10s}  {'Monitored':>10s}  {'Ratio':>6s}  Unit"
print(hdr)
print("-" * len(hdr))

for p in ['SO4','Ca','Cl','NO3','NH4','Cu','Ni','Zn','Co','Mo','As','Cr']:
    rate = frac_lev * LEV[p] + frac_kir * KIR[p]
    mass_leach = rate * prod

    # Correct unit: ug/L -> mg/L for trace metals before using in mass calc
    c_pit_mgl = SEP_2020[p] if SEP_UNITS[p] == 'mg/L' else SEP_2020[p] / 1000.0
    c_gb_mgl  = c_pit_mgl * 0.4   # rough Gruvberget fraction
    # mass in grams: c[mg/L] * V[L] / 1000[mg/g]
    V_pit_L = Q_pit * 1e9
    V_gb_L  = Q_gb  * 1e9
    V_out_L = Q_out * 1e9
    mass_pit = c_pit_mgl * V_pit_L / 1000.0
    mass_gb  = c_gb_mgl  * V_gb_L  / 1000.0
    total_g  = mass_leach + mass_pit + mass_gb

    c_new_mgl = total_g / V_out_L * 1000.0   # g/L -> mg/L
    c_pred = c_new_mgl if UNITS_MON[p] == 'mg/L' else c_new_mgl * 1000.0
    c_mon  = MON_2020[p]
    ratio  = c_pred / c_mon if c_mon > 0 else float('nan')

    flag = "OK" if 0.5 <= ratio <= 2.0 else ("HIGH" if ratio > 2 else "LOW")
    print(f"{p:4s}  {mass_leach:12.1f}  {mass_pit:12.1f}  {mass_gb:12.1f}  {c_pred:10.3f}  {c_mon:10.3f}  {ratio:6.2f}  {UNITS_MON[p]:5s}  {flag}")

print()
print("Diagnosis:")
print("- Ratio >> 1 means model over-predicts (rate too high or wrong unit in pit load)")
print("- Ratio << 1 means model under-predicts (missing load source or rate too low)")
print("- Cu/Ni/Zn/Co/As/Cr low -> leach rate swamped by pit water contribution")
print("- NH4 high -> Kiruna NH4 rate (3.047 g/ton) is from AMD leachate not drainage")
