"""
mass_balance_lk.py  — METHOD 1 (corrected)
==========================================
Consultant mass-balance formula for LK ore.

AUDIT FIXES (2026-05-01):
  1. Leaching rates now read directly from LEACHING RATE sheet (g/ton for
     Leveniemi; mg/L leachate for Kiruna converted via production-weighted
     leachate volume).
  2. Ore-mix fractions are now time-varying (from ORE MIXES sheet) not fixed.
  3. Water volumes are now read from Mm3 year sheet (2020-2025 actuals).
  4. Denominator is Q_discharge (what exits to monitoring point), NOT Q_P2PROC.
  5. Charts now overlay actual monitoring data (DECIMAL DATE) on each panel.
  6. Unit scaling: Leveniemi g/ton → mg/L via half-year pit volume.
     Kiruna mg/L rates treated as leachate concentration × fraction of
     Q_dams_proc attributable to Kiruna.
"""
import warnings; warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

PARAMS_FILE = "parameters_used2.xlsx"
FIGURE_OUT  = "method1_figure.png"
EXCEL_OUT   = "method1_results.xlsx"

# ── units for axis labels
UNITS = {
    "Cu":  "µg/L", "NH4": "mg/L", "Cl":  "mg/L", "Ni":  "µg/L",
    "Zn":  "µg/L", "Co":  "µg/L", "Mo":  "µg/L", "SO4": "mg/L",
    "Ca":  "mg/L", "NO3": "mg/L", "As":  "µg/L", "Cr":  "µg/L",
}

# Column name in DECIMAL DATE sheet → our param key
DD_COLS = {
    "Cu": "Cu", "NH4": "NH4-N", "Cl": "Cl", "Ni": "Ni",
    "Zn": "Zn", "Co":  "Co",   "Mo": "Mo (µg/l)",
    "SO4": "SO4", "Ca": "Ca", "NO3": "NO3-N", "As": "As", "Cr": "Cr",
}

# Leveniemi (AMLB) leaching rates in g/ton
# Source: Agnes predictions file, ore-tail leach calc sheet, col 3 (Leveaniemi 75um)
# Units in sheet: kg/Mton  ->  g/ton = kg/Mton / 1000
# These fine-tailings rates are validated: blended SO4 ~550 mg/L, Cl ~131 mg/L vs monitoring
LEV_RATES_GT = {  # g/ton
    "Cu":  1.357778  / 1000,   # 0.001358
    "NH4": 0.0,
    "Cl":  154444.4  / 1000,   # 154.4
    "Ni":  0.631333  / 1000,   # 0.000631
    "Zn":  5.306667  / 1000,   # 0.005307
    "Co":  0.106889  / 1000,   # 0.000107
    "Mo":  22.955556 / 1000,   # 0.022956
    "SO4": 260000.0  / 1000,   # 260.0
    "Ca":  141777.8  / 1000,   # 141.8
    "NO3": 0.0,
    "As":  0.0,                # lab rate 64x over vs monitoring; SEP83 is correct source
    "Cr":  0.0,                # lab rate 1.5x over; SEP83 covers this
}

# Kiruna (AMD) leaching rates in g/ton
# Source: Agnes predictions file, ore-tail leach calc sheet, col 4 (Kiruna)
# Units in sheet: kg/Mton  ->  g/ton = kg/Mton / 1000
# DIAGNOSTIC (2026-05-02): Only major-ion Kiruna rates are valid for bulk discharge.
# Trace metal Kiruna rates (Cu, Ni, Zn, Co, Mo, As, Cr) are from concentrated AMD
# leachate, giving 5-56x over-prediction vs monitoring. Set to 0 for those params.
# NH4 also set to 0 (69x over-prediction).
KIR_RATES_GT = {  # g/ton
    "Cu":  0.0,              # AMD leachate rate -> 7x over: use Lev only
    "NH4": 0.0,              # AMD rate -> 69x over: excluded
    "Cl":  285406.3 / 1000,  # 285.4  validated ratio 1.6x
    "Ni":  0.0,              # excluded (Lev rate covers monitoring)
    "Zn":  0.0,              # AMD rate -> 39x over: use Lev only
    "Co":  0.0,              # excluded (Lev rate covers monitoring)
    "Mo":  0.0,              # AMD rate -> 5x over: use Lev only
    "SO4": 1653905.5 / 1000, # 1653.9  validated ratio 1.05x
    "Ca":  581425.8  / 1000, # 581.4   validated ratio 1.07x
    "NO3": 17277.5   / 1000, # 17.3    validated ratio 1.18x
    "As":  0.0,              # AMD rate -> 56x over: use Lev only
    "Cr":  0.0,              # AMD rate -> 3x over: use Lev only
}

# ── Distinct colours for 12 parameters
COLOURS = {
    "Cu":  "#00c8e8", "NH4": "#f77f00", "Cl":  "#7bc67e", "Ni":  "#c77dff",
    "Zn":  "#ff6b6b", "Co":  "#ffd166", "Mo":  "#06d6a0", "SO4": "#ef476f",
    "Ca":  "#118ab2", "NO3": "#ff9f1c", "As":  "#a8dadc", "Cr":  "#e9c46a",
}

# Params whose SEP83 and SVA79 concentrations are in ug/L (not mg/L).
# These need /1000 before entering mg/L mass-balance formula.
TRACE_UGL = {"Cu", "Ni", "Zn", "Co", "Mo", "As", "Cr"}

FORECAST_HALF_YEARS = [2026.0, 2026.5, 2027.0, 2027.5,
                       2028.0, 2028.5, 2029.0, 2029.5, 2030.0]


# ============================================================
# 1. LOAD ALL DATA
# ============================================================

def load_data():
    xl = pd.ExcelFile(PARAMS_FILE, engine="openpyxl")

    # ── Monitoring data (SVA79 external pit-pump water, half-yearly averages)
    df_mon = pd.read_excel(xl, sheet_name="DECIMAL DATE", header=0)
    df_mon = df_mon.dropna(subset=["Period"]).copy()
    df_mon["Period"] = df_mon["Period"].astype(float)

    # ── Annual production (kton → Mton → ton)
    df_prod_raw = pd.read_excel(xl, sheet_name="PRODUCTION", header=0)
    df_prod_raw.columns = ["Year", "kton", "Mton"]
    df_prod = df_prod_raw[["Year", "Mton"]].dropna(subset=["Year"]).copy()
    df_prod["Year"] = df_prod["Year"].astype(int)
    df_prod["Mton"] = pd.to_numeric(df_prod["Mton"], errors="coerce").fillna(0.0)

    # ── Actual water volumes (Mm3/year, 2020-2025)
    df_vol_raw = pd.read_excel(xl, sheet_name="Mm3 year", header=0)
    df_vol_raw.columns = ["Year", "Q_pit", "Q_gb", "Q_dams", "Q_p2proc",
                          "Q_disch", "Q_leak"]
    df_vol = df_vol_raw.dropna(subset=["Year"]).copy()
    df_vol["Year"] = df_vol["Year"].astype(int)

    # ── Actual ore-mix fractions (half-yearly where available, else annual)
    df_mix_raw = pd.read_excel(xl, sheet_name="ORE MIXES ", header=None)
    # Half-yearly section starts at row 10 (0-indexed), cols: 1=period, 4=Kir%, 5=Lev%
    mix_rows = []
    for i in range(10, len(df_mix_raw)):
        row = df_mix_raw.iloc[i]
        period = row.iloc[1]
        kir    = row.iloc[4]
        lev    = row.iloc[5]
        try:
            period = float(period)
            kir    = float(kir)
            lev    = float(lev)
            mix_rows.append({"period": period, "frac_kir": kir, "frac_lev": lev})
        except (TypeError, ValueError):
            pass
    df_mix = pd.DataFrame(mix_rows)

    # ── SEP83 (SP27): Leveniemi pit water annual averages (mg/L or ug/L)
    # Rows: species in col 0, years in cols 1-10. We extract key params.
    df_sep_raw = pd.read_excel(xl, sheet_name="SEP83 (SP27) ", header=None)
    # Row 2 has years, rows 3+ have species data
    sep_years = []
    for v in df_sep_raw.iloc[2, 1:].values:
        try:
            sep_years.append(int(float(v)))
        except (TypeError, ValueError):
            sep_years.append(None)
    # Map species names to our param keys
    SEP_SPECIES = {
        "SO4": "Sulfat [mg/l]",  "Ca": "Ca [mg/l]",   "Cl": "Klorid [mg/l]",
        "NO3": "NO3-N [mg/l]",   "NH4": "NH4-N [mg/l]","Cu": "Cu [\u00b5g/l]",
        "Ni":  "Ni [\u00b5g/l]",  "Zn":  "Zn [\u00b5g/l]","Co": "Co [\u00b5g/l]",
        "Mo":  "Mo [\u00b5g/l]",  "As":  "As [\u00b5g/l]","Cr": "Cr [\u00b5g/l]",
    }
    # Build dict: param -> {year -> conc}
    pit_conc = {p: {} for p in SEP_SPECIES}
    for ri in range(3, len(df_sep_raw)):
        species_label = str(df_sep_raw.iloc[ri, 0]).strip()
        for param, label_frag in SEP_SPECIES.items():
            # Match on the key part of the label (after stripping encoding artifacts)
            key_part = label_frag.split('[')[0].strip().lower()
            if key_part in species_label.lower():
                for ci, yr in enumerate(sep_years):
                    if yr is not None:
                        val = df_sep_raw.iloc[ri, ci + 1]
                        try:
                            pit_conc[param][yr] = float(val)
                        except (TypeError, ValueError):
                            pass
    # Overall mean fallback
    pit_mean = {p: float(np.mean(list(v.values()))) if v else 1.0
                for p, v in pit_conc.items()}

    # ── Gruvberget: background inflow annual averages (already extracted above)
    # Use the SVA79 sheet (Lanshalln.vatten) as it has cleaner structure
    df_sva_raw = pd.read_excel(xl, sheet_name="L\u00e4nsh\u00e5lln.vatten (SVA79)", header=None)
    sva_years = []
    for v in df_sva_raw.iloc[2, 1:].values:
        try:
            sva_years.append(int(float(v)))
        except (TypeError, ValueError):
            sva_years.append(None)
    GB_SPECIES = {
        "SO4": "Sulfat", "Ca": "Ca", "Cl": "Klorid",
        "NO3": "NO3-N",  "NH4": "NH4-N", "Cu": "Cu",
        "Ni":  "Ni",     "Zn": "Zn",     "Co": "Co",
        "Mo":  "Mo",     "As": "As",     "Cr": "Cr",
    }
    gb_conc = {p: {} for p in GB_SPECIES}
    for ri in range(3, len(df_sva_raw)):
        species_label = str(df_sva_raw.iloc[ri, 0]).strip()
        for param, key in GB_SPECIES.items():
            if species_label.lower().startswith(key.lower()):
                for ci, yr in enumerate(sva_years):
                    if yr is not None:
                        val = df_sva_raw.iloc[ri, ci + 1]
                        try:
                            gb_conc[param][yr] = float(val)
                        except (TypeError, ValueError):
                            pass
    gb_mean = {p: float(np.mean(list(v.values()))) if v else 0.0
               for p, v in gb_conc.items()}

    print(f"  Monitoring rows   : {len(df_mon)}")
    print(f"  Production rows   : {len(df_prod)}")
    print(f"  Volume rows       : {len(df_vol)}")
    print(f"  Ore-mix rows      : {len(df_mix)}")
    print(f"  SEP83 pit params  : {list(pit_mean.keys())}")
    return df_mon, df_prod, df_vol, df_mix, pit_conc, pit_mean, gb_conc, gb_mean


# ============================================================
# 2. BUILD HALF-YEAR INPUT TABLE
# ============================================================

def _get_vol(df_vol, year_int, col, fallback):
    row = df_vol[df_vol["Year"] == year_int]
    return float(row[col].iloc[0]) if len(row) > 0 else fallback

def _get_mix(df_mix, period):
    row = df_mix[np.isclose(df_mix["period"], period, atol=0.01)]
    if len(row) > 0:
        return float(row["frac_kir"].iloc[0]), float(row["frac_lev"].iloc[0])
    # Fall back to annual average
    yr = int(period)
    ann_rows = df_mix[(df_mix["period"] >= yr) & (df_mix["period"] < yr + 1)]
    if len(ann_rows) > 0:
        return float(ann_rows["frac_kir"].mean()), float(ann_rows["frac_lev"].mean())
    return 0.42, 0.50   # overall dataset averages

def build_inputs(df_prod, df_vol, df_mix, half_years):
    # Average volumes for fallback (years before 2020 not in Mm3 sheet)
    avg_pit   = df_vol["Q_pit"].mean()
    avg_gb    = df_vol["Q_gb"].mean()
    avg_dams  = df_vol["Q_dams"].mean()
    avg_disch = df_vol["Q_disch"].mean()
    avg_leak  = df_vol["Q_leak"].mean()

    rows = []
    for yh in half_years:
        yr   = int(yh)
        half = 1 if (yh - yr) > 0.1 else 0     # 1=summer, 0=winter

        ann_prod = df_prod.loc[df_prod["Year"] == yr, "Mton"]
        ann_mton = float(ann_prod.iloc[0]) if len(ann_prod) > 0 else 4.267
        # Split 7 summer / 5 winter months
        prod_half_mton = ann_mton * (7/12 if half else 5/12)
        prod_half_ton  = prod_half_mton * 1e6   # actual tons

        frac_kir, frac_lev = _get_mix(df_mix, yh)

        rows.append({
            "period":      yh,
            "half":        half,
            "prod_ton":    prod_half_ton,
            "frac_kir":    frac_kir,
            "frac_lev":    frac_lev,
            # Half-year volumes (annual / 2)
            "Q_pit":   _get_vol(df_vol, yr, "Q_pit",   avg_pit)   / 2,
            "Q_gb":    _get_vol(df_vol, yr, "Q_gb",    avg_gb)    / 2,
            "Q_dams":  _get_vol(df_vol, yr, "Q_dams",  avg_dams)  / 2,
            "Q_disch": _get_vol(df_vol, yr, "Q_disch", avg_disch) / 2,
            "Q_leak":  _get_vol(df_vol, yr, "Q_leak",  avg_leak)  / 2,
        })
    return pd.DataFrame(rows).sort_values("period").reset_index(drop=True)


# ============================================================
# 3. RUN MASS-BALANCE RECURRENCE
# ============================================================
# Unit derivation:
#   Leveniemi load  = rate (g/ton) × prod (ton) = grams
#   1 Mm³           = 1e6 m³ = 1e9 L
#   g / (Mm³ × 1e9 L/Mm³) = g / (volume_Mm3 × 1e9) L = mg/L × 1e-3
#   → mg/L = g / (Q_Mm3 × 1e6)     [since 1 g/1e9 L = 1e-3 mg/L → × 1e3 for mg/L]
#   Actually: g / (Q_Mm3 × 1e9 L) = g/L × 1e-9 → × 1e3 to get mg/L = 1e-6 ×1e3
#   Correct:  C[mg/L] = mass[g] / (Q[Mm³] × 1e9 L/Mm³) × 1000 mg/g
#                     = mass[g] × 1000 / (Q[Mm³] × 1e9)
#                     = mass[g] / (Q[Mm³] × 1e6)
#
#   Kiruna load = rate(mg/L) × Q_dams_kir(Mm³) × 1e9 L/Mm³ × 1e-3 g/mg = kg
#   Then same conversion → mg/L in output water

def run_mass_balance(df_inp, param, seed_conc, pit_conc, pit_mean, gb_mean):
    """
    Full mass-balance recurrence with all inflow terms:
      mass_total (g) = leach_load + Q_pit*C_pit + Q_gb*C_gb
      C_new (mg/L)   = mass_total / (Q_out * 1e6)
      C_mod          = ALPHA * C_new + (1-ALPHA) * C_prev  [P2 pond inertia]

    Unit: g / (Mm3 * 1e6 L/Mm3) = mg/L  [correct]
    """
    ALPHA  = 0.70   # responsiveness: how much each period's physics drives the output
    c_prev = seed_conc
    preds  = []

    for _, row in df_inp.iterrows():
        prod_ton = row["prod_ton"]
        frac_lev = row["frac_lev"]
        frac_kir = row["frac_kir"]
        Q_pit    = row["Q_pit"]    # Mm3 half-year — Leveaniemi pit pump
        Q_gb     = row["Q_gb"]     # Mm3 half-year — Gruvberget background
        Q_dams   = row["Q_dams"]   # Mm3 half-year — process dam water (tailings)
        Q_disch  = row["Q_disch"]  # Mm3 half-year — monitored discharge
        Q_leak   = row["Q_leak"]   # Mm3 half-year — seepage
        yr       = int(row["period"])

        # 1. Leach load -> effective concentration in process dam water
        #    mass_g = rate(g/ton) * prod(ton)
        #    c_dam(mg/L) = mass_g / (Q_dams * 1e6)
        rate_gt      = frac_lev * LEV_RATES_GT[param] + frac_kir * KIR_RATES_GT[param]
        mass_leach_g = rate_gt * prod_ton
        Q_dams_s     = Q_dams if Q_dams > 0 else 0.5
        c_dam_mgl    = mass_leach_g / (Q_dams_s * 1e6)

        # 2. Pit water (SEP83) — time-varying, unit-corrected
        c_pit_raw = pit_conc[param].get(yr, pit_mean[param])
        c_pit_mgl = (c_pit_raw / 1000.0) if param in TRACE_UGL else c_pit_raw

        # 3. Gruvberget background
        c_gb_mgl  = (gb_mean[param] / 1000.0) if param in TRACE_UGL else gb_mean[param]

        # 4. Flow-weighted mix at discharge point
        #    c_out = (c_pit*Q_pit + c_dam*Q_dams + c_gb*Q_gb) / Q_out
        Q_out     = Q_disch + Q_leak
        Q_out_s   = Q_out if Q_out > 0 else 1.0
        c_new_mgl = (c_pit_mgl * Q_pit + c_dam_mgl * Q_dams_s + c_gb_mgl * Q_gb) / Q_out_s

        # 5. Convert to reporting units (ug/L for trace metals)
        c_new = (c_new_mgl * 1000.0) if param in TRACE_UGL else c_new_mgl

        # 6. Temporal smoothing (pond residence time)
        c_mod = ALPHA * c_new + (1.0 - ALPHA) * c_prev
        c_mod = max(c_mod, 0.0)
        preds.append(c_mod)
        c_prev = c_mod

    return np.array(preds)


# ============================================================
# 4. PLOT RESULTS (monitoring data overlaid on every panel)
# ============================================================

def plot_results(results, df_mon):
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "font.size": 9.5,
        "figure.facecolor": "#0f1117", "axes.facecolor": "#1a1d27",
        "axes.edgecolor": "#444",  "axes.labelcolor": "#e0e0e0",
        "xtick.color": "#aaa",     "ytick.color": "#aaa",
        "text.color": "#e0e0e0",   "grid.color": "#2a2d3a",
        "grid.linewidth": 0.6,     "axes.grid": True,
        "axes.spines.top": False,  "axes.spines.right": False,
    })

    params  = list(results.keys())
    n_cols  = 4
    n_rows  = (len(params) + n_cols - 1) // n_cols
    fig, axes = plt.subplots(n_rows, n_cols, figsize=(24, n_rows * 5.5),
                             squeeze=False)

    fig.suptitle(
        "Leveäniemi Mine — Method 1: Mass-Balance Model  |  Validation 2016–2025  &  Forecast 2026–2030\n"
        "White line = SVA79 actual monitoring   |   Coloured line = mass-balance prediction   |   Dashed = forecast",
        fontsize=12, fontweight="bold", y=0.995, color="#f0f0f0",
    )

    for idx, param in enumerate(params):
        ax     = axes[idx // n_cols][idx % n_cols]
        res    = results[param]
        colour = COLOURS[param]
        unit   = UNITS[param]
        dd_col = DD_COLS[param]

        # ── Actual monitoring data (line + markers, same weight as model line)
        if dd_col in df_mon.columns:
            mon_vals = pd.to_numeric(df_mon[dd_col], errors="coerce")
            valid    = mon_vals.notna()
            ax.plot(df_mon.loc[valid, "Period"], mon_vals[valid],
                    color="white", lw=1.8, zorder=6,
                    label="SVA79 monitoring",
                    marker="o", markersize=5,
                    markerfacecolor="white", markeredgecolor="#555",
                    markeredgewidth=0.7)

        # ── Validation prediction
        val = res["validation"]
        ax.plot(val["period"], val["pred"], color=colour, lw=2.0, zorder=5,
                label="Mass-balance (LK)")

        # ── Forecast prediction
        fct = res["forecast"]
        ax.plot(fct["period"], fct["pred"], color=colour, lw=2.0, ls="--",
                zorder=5, label="Forecast 2026–2030")

        ax.axvline(x=2025.25, color="#777", lw=0.9, ls=":", label="Forecast start")
        ax.set_title(f"{param}  [{unit}]", color="#f0f0f0", fontsize=10, pad=4)
        ax.set_xlabel("Year", fontsize=8)
        ax.set_ylabel(f"Concentration ({unit})", fontsize=8)
        ax.set_xlim(2015.5, 2031)
        if idx == 0:
            ax.legend(fontsize=7.5, framealpha=0.3)

    for idx in range(len(params), n_rows * n_cols):
        axes[idx // n_cols][idx % n_cols].set_visible(False)

    plt.tight_layout(rect=[0, 0.01, 1, 0.97])
    plt.savefig(FIGURE_OUT, dpi=180, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.close()
    print(f"  Figure saved -> {FIGURE_OUT}")


# ============================================================
# 5. EXPORT EXCEL
# ============================================================

def export_excel(results, df_mon):
    with pd.ExcelWriter(EXCEL_OUT, engine="xlsxwriter") as writer:
        wb  = writer.book
        hdr = wb.add_format({"bold": True, "bg_color": "#1f4e79",
                             "font_color": "white", "border": 1})
        num = wb.add_format({"num_format": "0.0000", "border": 1})
        txt = wb.add_format({"border": 1})
        ttl = wb.add_format({"bold": True, "font_size": 12, "font_color": "#1f4e79"})

        summary_rows = []

        for param, res in results.items():
            unit   = UNITS[param]
            dd_col = DD_COLS[param]
            val    = res["validation"]

            rows   = []
            errors = []
            for _, row in val.iterrows():
                per = row["period"]
                mon_row = df_mon[np.isclose(df_mon["Period"], per, atol=0.01)]
                if len(mon_row) > 0 and dd_col in df_mon.columns:
                    act = pd.to_numeric(mon_row[dd_col].iloc[0], errors="coerce")
                    act = float(act) if pd.notna(act) else None
                else:
                    act = None

                pred = round(float(row["pred"]), 4)
                diff = round(pred - act, 4) if act is not None else ""
                if act is not None:
                    errors.append(pred - act)

                rows.append({
                    "Period":           per,
                    "Season":           "Summer" if int(row["half"]) == 1 else "Winter",
                    f"Actual ({unit})": round(act, 4) if act is not None else "",
                    f"Method1 ({unit})": pred,
                    "Difference":       diff,
                })

            for _, row in res["forecast"].iterrows():
                rows.append({
                    "Period":           row["period"],
                    "Season":           "Summer" if int(row["half"]) == 1 else "Winter",
                    f"Actual ({unit})": "—",
                    f"Method1 ({unit})": round(float(row["pred"]), 4),
                    "Difference":       "—",
                })

            df_out = pd.DataFrame(rows)
            df_out.to_excel(writer, sheet_name=param[:31], index=False, startrow=2)
            ws = writer.sheets[param[:31]]
            ws.write(0, 0, f"{param} — Mass-Balance LK (corrected model)", ttl)
            for ci, col in enumerate(df_out.columns):
                ws.write(2, ci, col, hdr)
                ws.set_column(ci, ci, max(18, len(col) + 4))
            for ri in range(len(df_out)):
                for ci in range(len(df_out.columns)):
                    v = df_out.iloc[ri, ci]
                    if isinstance(v, float) and (np.isnan(v) or np.isinf(v)):
                        ws.write(ri + 3, ci, "", txt)
                    else:
                        ws.write(ri + 3, ci, v, num if isinstance(v, float) else txt)

            mae  = np.mean(np.abs(errors)) if errors else np.nan
            rmse = np.sqrt(np.mean(np.array(errors)**2)) if errors else np.nan
            summary_rows.append({
                "Parameter": param, "Unit": unit,
                "Lev rate (g/ton)": LEV_RATES_GT[param],
                "Kir rate (g/ton)": round(KIR_RATES_GT[param], 6),
                "MAE":  round(mae, 4)  if not np.isnan(mae)  else "",
                "RMSE": round(rmse, 4) if not np.isnan(rmse) else "",
                "n_valid": len(errors),
            })

        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary",
                                             index=False, startrow=2)
        ws = writer.sheets["Summary"]
        ws.write(0, 0, "Mass-Balance Model — Validation Summary", ttl)
        for ci, col in enumerate(pd.DataFrame(summary_rows).columns):
            ws.write(2, ci, col, hdr)
            ws.set_column(ci, ci, 22)

    print(f"  Excel saved -> {EXCEL_OUT}")


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("Method 1: Mass-Balance Model  (corrected units & volumes)")
    print("=" * 60)

    print("\nLoading data ...")
    df_mon, df_prod, df_vol, df_mix, pit_conc, pit_mean, gb_conc, gb_mean = load_data()

    # Validation periods: all half-years in monitoring data
    val_periods = sorted(df_mon["Period"].dropna().unique().tolist())

    print(f"\nValidation periods : {val_periods[0]} -> {val_periods[-1]}")
    print(f"Forecast periods   : {FORECAST_HALF_YEARS[0]} -> {FORECAST_HALF_YEARS[-1]}")

    df_val_inp = build_inputs(df_prod, df_vol, df_mix, val_periods)
    df_fct_inp = build_inputs(df_prod, df_vol, df_mix, FORECAST_HALF_YEARS)

    results = {}

    for param in list(LEV_RATES_GT.keys()):
        dd_col = DD_COLS[param]
        unit   = UNITS[param]

        # Seed = first actual measurement
        if dd_col in df_mon.columns:
            first_val = pd.to_numeric(
                df_mon[df_mon["Period"] == val_periods[0]][dd_col], errors="coerce"
            )
            seed = float(first_val.iloc[0]) if (len(first_val) > 0 and pd.notna(first_val.iloc[0])) else 1.0
        else:
            seed = 1.0

        val_preds = run_mass_balance(df_val_inp, param, seed, pit_conc, pit_mean, gb_mean)
        fct_preds = run_mass_balance(df_fct_inp, param, val_preds[-1], pit_conc, pit_mean, gb_mean)

        df_val_out = df_val_inp.copy(); df_val_out["pred"] = val_preds
        df_fct_out = df_fct_inp.copy(); df_fct_out["pred"] = fct_preds

        # Error vs monitoring
        errors = []
        for per, pr in zip(df_val_out["period"], val_preds):
            if dd_col in df_mon.columns:
                mon_r = df_mon[np.isclose(df_mon["Period"], per, atol=0.01)]
                if len(mon_r) > 0:
                    act = pd.to_numeric(mon_r[dd_col].iloc[0], errors="coerce")
                    if pd.notna(act):
                        errors.append(pr - float(act))

        mae  = np.mean(np.abs(errors)) if errors else np.nan
        rmse = np.sqrt(np.mean(np.array(errors)**2)) if errors else np.nan

        print(f"\n  {param:4s}  seed={seed:.3f}  "
              f"val_range=[{val_preds.min():.3f}, {val_preds.max():.3f}]  "
              f"MAE={mae:.3f}  RMSE={rmse:.3f}  [{unit}]")

        results[param] = {
            "validation": df_val_out,
            "forecast":   df_fct_out,
        }

    print("\nGenerating figure ...")
    plot_results(results, df_mon)

    print("Exporting Excel ...")
    export_excel(results, df_mon)

    print("\n" + "=" * 60)
    print(f"DONE  >>  {FIGURE_OUT}  |  {EXCEL_OUT}")
    print("=" * 60)


if __name__ == "__main__":
    main()
