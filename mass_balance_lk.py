"""
mass_balance_lk.py
==================
Method 1: Consultant's mass-balance formula applied to LK (Leveaniemi-Kiruna) ore.

Steps:
  1. Load actual LK leaching rates from parameters_used.xlsx
  2. Load actual production volumes (2016-2025)
  3. Run the consultant's recurrence formula for 2016-2025 using LK inputs
  4. Compare predictions to actual measured concentrations (DECIMAL DATE sheet)
  5. Forecast 2026-2030
  6. Export comparison figure + Excel

Run:  python mass_balance_lk.py
"""
import warnings; warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

# ============================================================
# CONFIGURATION
# ============================================================

PARAMS_FILE   = "parameters_used.xlsx"
CONSUL_FILE   = "Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx"
FIGURE_OUT    = "method1_figure.png"
EXCEL_OUT     = "method1_results.xlsx"

# Default LK ore mix fraction (Leveaniemi / Kiruna)
# 50/50 default — user can change this
LK_FRAC_LEV = 0.50
LK_FRAC_KIR = 0.50

# ── Leveaniemi leaching rates (g/ton from LEACHING RATE sheet, row 25)
# Converting g/ton → kg/Mton: multiply by 1000
LEV_RATES_G_PER_TON = {
    "Cu":   0.004638,   # col 21 in row 25
    "NH4":  0.327503,   # col 3 (NH4-N) — g/ton
    "Cl":   41.5637,    # col 8
    "Ni":   0.001221,   # col 25
}
LEV_RATES = {k: v * 1000 for k, v in LEV_RATES_G_PER_TON.items()}  # now kg/Mton

# ── Kiruna leaching rates (consultant's established values, kg/Mton)
# Cu and NH4 not in the LEACHING RATE sheet → use consultant's ore-tail values
KIR_RATES = {
    "Cu":   11.668,     # from ore-tail leach calc sheet
    "NH4":  3047.3,     # from ore-tail leach calc sheet
    "Cl":   285406.0,   # from ore-tail leach calc sheet
    "Ni":   1.174,      # from ore-tail leach calc sheet
}

# ── LK blended rate = frac_lev * lev + frac_kir * kir
LK_RATES = {
    p: LK_FRAC_LEV * LEV_RATES[p] + LK_FRAC_KIR * KIR_RATES[p]
    for p in ["Cu", "NH4", "Cl", "Ni"]
}

# ── System constants from consultant's model (fixed volumes, Mm3)
# From NH4 block row 81: col20=25.21 (P2 to process), col22=21.42,
#   col24=0.11 (Tailings), col26=5.8 (Discharge), col28=1.7 (Leakage)
#   col14=1.26 (Gruvberget), col16=1.6 (Surface water summer)
Q_LOSS    = 0.36    # process plant loss (Mm3)
Q_SURF_S  = 1.6     # surface water inflow summer (Mm3)
Q_SURF_W  = 0.0     # surface water inflow winter
Q_TAIL    = 0.11    # tailings drainage (Mm3)
Q_GB      = 1.26    # Gruvberget inflow (Mm3)
Q_DISCH   = 5.8     # discharge (Mm3)
Q_LEAKAGE = 1.7     # leakage (Mm3)
Q_P2PROC  = 25.21   # P2 to process (Mm3) -- main storage volume
Q_TAIL_ST = 1.84    # Tailings storage (Mm3)

# ── Background concentrations (mg/L)
# From consultant rows for NH4: ditch=0.01, Gruvberget=1.0, surface=0.0
# Approximate C_pit from SVA79 monitoring data
C_SURF  = {"Cu": 0.0, "NH4": 0.0,   "Cl": 18.0, "Ni": 0.0}
C_TAIL  = {"Cu": 0.0, "NH4": 0.2,   "Cl": 0.0,  "Ni": 0.2}
C_GB    = {"Cu": 3.0, "NH4": 1.0,   "Cl": 18.0, "Ni": 0.15}
C_DITCH = {"Cu": 0.0, "NH4": 0.01,  "Cl": 0.01, "Ni": 0.0}
C_LEAK  = {"Cu": 0.0, "NH4": 0.2,   "Cl": 0.0,  "Ni": 0.2}

# ── Units for display
UNITS = {"Cu": "ug/L", "NH4": "mg/L", "Cl": "mg/L", "Ni": "ug/L"}
COLOURS = {"Cu": "#00c8e8", "NH4": "#f77f00", "Cl": "#7bc67e", "Ni": "#c77dff"}

FORECAST_YEARS_HALF = [2026.0, 2026.5, 2027.0, 2027.5, 2028.0,
                       2028.5, 2029.0, 2029.5, 2030.0]
# Assumed future production (Mton/year) — use 2025 as baseline
FUTURE_PROD_ANNUAL = 4.267  # Mton, same as 2025


# ============================================================
# 1. LOAD DATA
# ============================================================

def load_data():
    xl = pd.ExcelFile(PARAMS_FILE, engine="openpyxl")

    # Actual measured concentrations by half-year
    df_actual = pd.read_excel(xl, sheet_name="DECIMAL DATE", header=0)
    df_actual = df_actual[["Period", "Season", "Cu", "NH4-N", "Cl", "Ni"]].copy()
    df_actual = df_actual.rename(columns={"NH4-N": "NH4"})
    df_actual = df_actual.dropna(subset=["Period"])

    # Actual production volumes (kton → Mton)
    df_prod_raw = pd.read_excel(xl, sheet_name="PRODUCTION", header=0)
    df_prod_raw.columns = ["Year", "kton", "Mton"]
    df_prod = df_prod_raw[["Year", "Mton"]].dropna(subset=["Year"]).copy()
    df_prod["Year"] = df_prod["Year"].astype(int)
    df_prod["Mton"] = df_prod["Mton"].fillna(0.0)

    # Pit pump volumes from first sheet (Leveaniemi column = pit pump)
    df_vol_raw = pd.read_excel(xl, sheet_name=xl.sheet_names[0], header=None)
    vol_data = []
    for i in range(df_vol_raw.shape[0]):
        row = df_vol_raw.iloc[i]
        yr_val = row.iloc[0]
        lev_val = row.iloc[1]
        if pd.notna(yr_val) and pd.notna(lev_val):
            try:
                vol_data.append({"Year": float(yr_val), "Q_pit": float(lev_val)})
            except:
                pass
    df_vol = pd.DataFrame(vol_data)

    print("Actual data rows:", len(df_actual))
    print("Production rows:", len(df_prod))
    print("Pit pump rows:", len(df_vol))
    return df_actual, df_prod, df_vol


# ============================================================
# 2. BUILD HALF-YEAR INPUT TABLE
# ============================================================

def build_inputs(df_prod, df_vol, years_half):
    """
    For each half-year period, compute:
      - prod_lk  : LK production (Mton)
      - leach_lk : LK leaching load per parameter (kg), before rate applied
      - q_pit    : pit pump volume (Mm3) for that period
    """
    rows = []
    for yh in years_half:
        yr   = int(yh)
        half = 1 if (yh - yr) > 0.1 else 0   # 1=summer, 0=winter

        # Production: split annual total by season (7 months summer / 5 months winter)
        ann_prod = df_prod.loc[df_prod["Year"] == yr, "Mton"]
        if len(ann_prod) > 0:
            ann = float(ann_prod.iloc[0])
        else:
            ann = FUTURE_PROD_ANNUAL  # use future projection
        prod_lk = ann * (7/12) if half == 1 else ann * (5/12)

        # Pit pump volume for this period
        pit_row = df_vol[df_vol["Year"] == yh]
        if len(pit_row) > 0:
            q_pit = float(pit_row["Q_pit"].iloc[0])
        else:
            q_pit = 1.5  # fallback average

        rows.append({
            "year":    yh,
            "half":    half,
            "prod_lk": prod_lk,
            "q_pit":   q_pit,
        })

    return pd.DataFrame(rows).sort_values("year").reset_index(drop=True)


# ============================================================
# 3. RUN THE CONSULTANT'S MASS-BALANCE FORMULA
# ============================================================

def run_mass_balance(df_inputs, param, seed_conc, c_pit):
    """
    Consultant's recurrence formula applied to LK ore.

    All volumes in Mm3, leaching loads in kg → concentration in mg/L
    (1 kg / 1 Mm3 = 1 kg / 10^6 m3 = 10^-6 kg/L = 10^-3 g/L = 1 mg/L)

    From the consultant's formula (NH4 block, row 81 structure):
      Gains = leach_lk + Q_pit*C_pit + Q_GB*C_GB + Q_SURF*C_SURF
              + Q_TAIL*C_TAIL + Q_LEAK*C_LEAK - Q_LOSS*C_prev
      Total_vol = Q_pit + Q_GB + Q_SURF + Q_TAIL + Q_LEAKAGE
                  + Q_P2PROC + Q_TAIL_ST - Q_LOSS - Q_DISCH
      C_active = Gains / Total_vol   [mg/L]
      C_mod(t) = (Gains * Q_active + Q_TAIL_ST * C_prev) / (Q_active + Q_TAIL_ST)
    """
    lk_rate = LK_RATES[param]
    c_gb    = C_GB[param]
    c_tail  = C_TAIL[param]
    c_ditch = C_DITCH[param]
    c_leak  = C_LEAK[param]

    preds  = []
    c_prev = seed_conc

    for _, row in df_inputs.iterrows():
        prod_lk = row["prod_lk"]
        q_pit   = row["q_pit"]
        half    = int(row["half"])
        q_surf  = Q_SURF_S if half == 1 else Q_SURF_W

        # Total leaching mass (kg) from LK ore
        leach_kg = prod_lk * lk_rate

        # Total gains (kg equivalent, using Mm3 volumes)
        gains = (leach_kg
                 + q_pit   * c_pit
                 + Q_GB    * c_gb
                 + q_surf  * C_SURF[param]
                 + Q_TAIL  * c_tail
                 + Q_LEAKAGE * c_leak
                 - Q_LOSS  * c_prev)

        # Total active water volume (Mm3)
        q_active = (q_pit + Q_GB + q_surf + Q_TAIL + Q_LEAKAGE
                    + Q_P2PROC - Q_LOSS - Q_DISCH)

        # Active concentration (mg/L = kg/Mm3)
        c_active = gains / q_active if q_active > 0 else 0.0

        # Storage-weighted recurrence
        c_mod = ((c_active * q_active + Q_TAIL_ST * c_prev)
                 / (q_active + Q_TAIL_ST))

        c_mod  = max(c_mod, 0.0)
        preds.append(c_mod)
        c_prev = c_mod

    return np.array(preds)


# ============================================================
# 4. PLOT RESULTS
# ============================================================

def plot_results(results):
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "font.size": 11,
        "figure.facecolor": "#0f1117", "axes.facecolor": "#1a1d27",
        "axes.edgecolor": "#444", "axes.labelcolor": "#e0e0e0",
        "xtick.color": "#aaa", "ytick.color": "#aaa",
        "text.color": "#e0e0e0", "grid.color": "#2a2d3a",
        "grid.linewidth": 0.6, "axes.grid": True,
        "axes.spines.top": False, "axes.spines.right": False,
    })

    params = list(results.keys())
    fig, axes = plt.subplots(2, 2, figsize=(16, 11))
    axes = axes.flatten()

    fig.suptitle(
        "Leveaniemi Mine — Method 1: Mass-Balance Model Validation & Forecast\n"
        f"LK Ore ({int(LK_FRAC_LEV*100)}% Leveaniemi / {int(LK_FRAC_KIR*100)}% Kiruna)",
        fontsize=14, fontweight="bold", y=0.98, color="#f0f0f0",
    )

    for ax, param in zip(axes, params):
        res    = results[param]
        colour = COLOURS[param]
        unit   = UNITS[param]

        # Actual measured data
        act   = res["actual"]
        ax.scatter(act["Period"], act[param], color="#ffffff", s=60, zorder=5,
                   label="Actual measured", marker="o", edgecolors="#888", linewidths=0.8)

        # Method 1 validation (2016-2025)
        val = res["validation"]
        ax.plot(val["year"], val["pred"], color=colour, lw=2.2,
                label="Method 1: Mass-balance (LK)")

        # Method 1 forecast (2026-2030)
        fct = res["forecast"]
        ax.plot(fct["year"], fct["pred"], color=colour, lw=2.2,
                ls="--", label="Method 1: Forecast 2026-2030")

        # Vertical boundary line
        ax.axvline(x=2025.75, color="#666", lw=0.9, ls=":")
        ax.text(2025.8, ax.get_ylim()[1] * 0.95, "Forecast >>",
                fontsize=8, color="#888")

        ax.set_title(f"{param}  [{unit}]", color="#f0f0f0", pad=6)
        ax.set_xlabel("Year")
        ax.set_ylabel(f"Concentration  ({unit})")
        ax.set_xlim(2015.5, 2031)
        ax.legend(fontsize=8.5)

    plt.tight_layout(rect=[0, 0.04, 1, 0.95])
    plt.savefig(FIGURE_OUT, dpi=180, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.show()
    print(f"Figure saved: {FIGURE_OUT}")


# ============================================================
# 5. EXPORT EXCEL
# ============================================================

def export_excel(results):
    with pd.ExcelWriter(EXCEL_OUT, engine="xlsxwriter") as writer:
        wb  = writer.book
        hdr = wb.add_format({"bold": True, "bg_color": "#1f4e79",
                             "font_color": "white", "border": 1})
        num = wb.add_format({"num_format": "0.0000", "border": 1})
        txt = wb.add_format({"border": 1})
        ttl = wb.add_format({"bold": True, "font_size": 12,
                             "font_color": "#1f4e79"})

        for param, res in results.items():
            unit = UNITS[param]
            val  = res["validation"]
            act  = res["actual"].set_index("Period")

            rows = []
            for _, row in val.iterrows():
                yr  = row["year"]
                act_val = act.loc[yr, param] if yr in act.index else np.nan
                rows.append({
                    "Year":                yr,
                    "Season":              "Summer" if row["half"] == 1 else "Winter",
                    f"Actual ({unit})":    round(act_val, 4) if pd.notna(act_val) else "",
                    f"Method1 ({unit})":   round(row["pred"], 4),
                    "Difference":          round(row["pred"] - act_val, 4)
                                           if pd.notna(act_val) else "",
                })
            # Add forecast rows
            for _, row in res["forecast"].iterrows():
                rows.append({
                    "Year":               row["year"],
                    "Season":             "Summer" if row["half"] == 1 else "Winter",
                    f"Actual ({unit})":   "—",
                    f"Method1 ({unit})":  round(row["pred"], 4),
                    "Difference":         "—",
                })

            df_out = pd.DataFrame(rows)
            df_out.to_excel(writer, sheet_name=param, index=False, startrow=2)
            ws = writer.sheets[param]
            ws.write(0, 0, f"{param} — Mass-Balance LK Model  |  "
                           f"Mix: {int(LK_FRAC_LEV*100)}% Lev / {int(LK_FRAC_KIR*100)}% Kir", ttl)
            for ci, col in enumerate(df_out.columns):
                ws.write(2, ci, col, hdr)
                ws.set_column(ci, ci, max(18, len(col) + 4))
            for ri in range(len(df_out)):
                for ci in range(len(df_out.columns)):
                    v = df_out.iloc[ri, ci]
                    ws.write(ri + 3, ci, v,
                             num if isinstance(v, float) else txt)

        # Summary sheet
        summary_rows = []
        for param, res in results.items():
            act = res["actual"].set_index("Period")
            preds = res["validation"]
            errors = []
            for _, row in preds.iterrows():
                yr = row["year"]
                if yr in act.index and pd.notna(act.loc[yr, param]):
                    errors.append(row["pred"] - act.loc[yr, param])
            mae  = np.mean(np.abs(errors)) if errors else np.nan
            rmse = np.sqrt(np.mean(np.array(errors)**2)) if errors else np.nan
            summary_rows.append({
                "Parameter":           param,
                "Unit":                UNITS[param],
                "Mix":                 f"{int(LK_FRAC_LEV*100)}% Lev / {int(LK_FRAC_KIR*100)}% Kir",
                "LK leach rate":       round(LK_RATES[param], 4),
                "MAE":                 round(mae, 4) if not np.isnan(mae) else "",
                "RMSE":                round(rmse, 4) if not np.isnan(rmse) else "",
                "Validation periods":  len(errors),
            })
        df_sum = pd.DataFrame(summary_rows)
        df_sum.to_excel(writer, sheet_name="Summary", index=False, startrow=2)
        ws = writer.sheets["Summary"]
        ws.write(0, 0, "Mass-Balance Model Validation Summary", ttl)
        for ci, col in enumerate(df_sum.columns):
            ws.write(2, ci, col, hdr)
            ws.set_column(ci, ci, max(20, len(col) + 4))

    print(f"Excel saved: {EXCEL_OUT}")


# ============================================================
# MAIN
# ============================================================

def main():
    print("="*60)
    print("Method 1: Mass-Balance Model for LK Ore")
    print(f"Mix: {int(LK_FRAC_LEV*100)}% Leveaniemi / {int(LK_FRAC_KIR*100)}% Kiruna")
    print("="*60)

    print("\nLK Leaching Rates (kg/Mton):")
    for p, r in LK_RATES.items():
        print(f"  {p}: {r:.4f}  "
              f"(Lev={LEV_RATES[p]:.4f}, Kir={KIR_RATES[p]:.4f})")

    print("\nLoading data ...")
    df_actual, df_prod, df_vol = load_data()

    # Validation period: all half-years with actual data
    val_years  = sorted(df_actual["Period"].dropna().tolist())
    df_val_inp = build_inputs(df_prod, df_vol, val_years)

    # Forecast period: 2026-2030
    df_fct_inp = build_inputs(df_prod, df_vol, FORECAST_YEARS_HALF)

    results = {}

    for param in ["Cu", "NH4", "Cl", "Ni"]:
        print(f"\n  Running {param} ...")

        # Seed concentration: first actual measurement
        act_col = param
        first_actual = df_actual[df_actual["Period"] == val_years[0]][act_col]
        seed = float(first_actual.iloc[0]) if len(first_actual) > 0 else 1.0

        # Approximate pit water concentration from actual data averages
        c_pit = float(df_actual[act_col].mean()) * 0.3   # pit water is ~30% of surface conc

        # Run validation
        val_preds = run_mass_balance(df_val_inp, param, seed, c_pit)
        df_val_out = df_val_inp.copy()
        df_val_out["pred"] = val_preds

        # Seed forecast from last validation prediction
        seed_fct   = val_preds[-1]
        fct_preds  = run_mass_balance(df_fct_inp, param, seed_fct, c_pit)
        df_fct_out = df_fct_inp.copy()
        df_fct_out["pred"] = fct_preds

        # Compute error vs actual
        act_sub = df_actual[["Period", act_col]].rename(
            columns={act_col: param}).set_index("Period")
        errors = []
        for yr, p in zip(df_val_out["year"], val_preds):
            if yr in act_sub.index and pd.notna(act_sub.loc[yr, param]):
                errors.append(p - act_sub.loc[yr, param])

        mae  = np.mean(np.abs(errors)) if errors else np.nan
        rmse = np.sqrt(np.mean(np.array(errors)**2)) if errors else np.nan
        print(f"    Validation MAE:  {mae:.4f}  {UNITS[param]}")
        print(f"    Validation RMSE: {rmse:.4f}  {UNITS[param]}")
        print(f"    Forecast 2026-2030 range: "
              f"{fct_preds.min():.4f} -- {fct_preds.max():.4f}  {UNITS[param]}")

        results[param] = {
            "actual":     df_actual[["Period", "Season", param]].copy()
                          if param in df_actual.columns
                          else df_actual[["Period", "Season",
                                          "Cu", "NH4", "Cl", "Ni"]].rename(
                              columns={"NH4-N": "NH4"}),
            "validation": df_val_out,
            "forecast":   df_fct_out,
        }

    print("\nGenerating figure ...")
    plot_results(results)

    print("Exporting Excel ...")
    export_excel(results)

    print("\n" + "="*60)
    print("DONE — Method 1 complete")
    print(f"  >> {FIGURE_OUT}")
    print(f"  >> {EXCEL_OUT}")
    print("="*60)


if __name__ == "__main__":
    main()
