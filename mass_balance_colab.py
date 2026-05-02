# ============================================================
# Leveäniemi Water Quality — Mass-Balance Model (Colab Version)
# ============================================================
# HOW TO USE IN GOOGLE COLAB:
#   1. Upload this file to Colab (or paste each section into cells)
#   2. Run Cell 1 to install packages
#   3. Run Cell 2 to upload your Excel file
#   4. Run Cell 3+ to execute the model
# ============================================================

# ── CELL 1: Install packages ─────────────────────────────────
# Paste this into the first Colab cell and run it:
#
#   !pip install openpyxl xlsxwriter --quiet

# ── CELL 2: Upload data file ──────────────────────────────────
# Paste this into the second Colab cell and run it.
# A file picker will appear — upload parameters_used2.xlsx
#
#   from google.colab import files
#   uploaded = files.upload()   # upload parameters_used2.xlsx
#   PARAMS_FILE = list(uploaded.keys())[0]
#   print("Using file:", PARAMS_FILE)

# ── CELL 3: Full model (paste everything below) ───────────────

import warnings; warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt

# Detect Colab environment
try:
    from google.colab import files as colab_files
    IN_COLAB = True
except ImportError:
    colab_files = None
    IN_COLAB = False

# PARAMS_FILE is set by the upload cell in Colab.
# When running locally, set it here:
# PARAMS_FILE = "parameters_used2.xlsx"

FIGURE_OUT = "method1_figure.png"
EXCEL_OUT  = "method1_results.xlsx"

UNITS = {
    "Cu": "µg/L", "NH4": "mg/L", "Cl":  "mg/L", "Ni":  "µg/L",
    "Zn": "µg/L", "Co":  "µg/L", "Mo":  "µg/L", "SO4": "mg/L",
    "Ca": "mg/L", "NO3": "mg/L", "As":  "µg/L", "Cr":  "µg/L",
}

DD_COLS = {
    "Cu": "Cu", "NH4": "NH4-N", "Cl": "Cl", "Ni": "Ni",
    "Zn": "Zn", "Co":  "Co",   "Mo": "Mo (µg/l)",
    "SO4": "SO4", "Ca": "Ca", "NO3": "NO3-N", "As": "As", "Cr": "Cr",
}

# Leveäniemi fine-tailings leaching rates (Agnes ore-tail leach calc, col 3)
# Units: kg/Mton in source → divide by 1000 → g/ton
LEV_RATES_GT = {
    "Cu":  1.357778  / 1000,
    "NH4": 0.0,
    "Cl":  154444.4  / 1000,
    "Ni":  0.631333  / 1000,
    "Zn":  5.306667  / 1000,
    "Co":  0.106889  / 1000,
    "Mo":  22.955556 / 1000,
    "SO4": 260000.0  / 1000,
    "Ca":  141777.8  / 1000,
    "NO3": 0.0,
    "As":  12.511111 / 1000,
    "Cr":  0.163111  / 1000,
}

# Kiruna rates — only major ions validated; trace metals excluded (5–69x over-prediction)
KIR_RATES_GT = {
    "Cu":  0.0,
    "NH4": 0.0,
    "Cl":  285406.3  / 1000,
    "Ni":  0.0,
    "Zn":  0.0,
    "Co":  0.0,
    "Mo":  0.0,
    "SO4": 1653905.5 / 1000,
    "Ca":  581425.8  / 1000,
    "NO3": 17277.5   / 1000,
    "As":  0.0,
    "Cr":  0.0,
}

COLOURS = {
    "Cu": "#00c8e8", "NH4": "#f77f00", "Cl":  "#7bc67e", "Ni":  "#c77dff",
    "Zn": "#ff6b6b", "Co":  "#ffd166", "Mo":  "#06d6a0", "SO4": "#ef476f",
    "Ca": "#118ab2", "NO3": "#ff9f1c", "As":  "#a8dadc", "Cr":  "#e9c46a",
}

# Trace metals stored in µg/L in SEP83/SVA79 — need ÷1000 before mg/L mass calc
TRACE_UGL = {"Cu", "Ni", "Zn", "Co", "Mo", "As", "Cr"}

FORECAST_HALF_YEARS = [2026.0, 2026.5, 2027.0, 2027.5,
                       2028.0, 2028.5, 2029.0, 2029.5, 2030.0]


# ── DATA LOADING ─────────────────────────────────────────────

def load_data(params_file):
    xl = pd.ExcelFile(params_file, engine="openpyxl")

    df_mon = pd.read_excel(xl, sheet_name="DECIMAL DATE", header=0)
    df_mon = df_mon.dropna(subset=["Period"]).copy()
    df_mon["Period"] = df_mon["Period"].astype(float)

    df_prod_raw = pd.read_excel(xl, sheet_name="PRODUCTION", header=0)
    df_prod_raw.columns = ["Year", "kton", "Mton"]
    df_prod = df_prod_raw[["Year", "Mton"]].dropna(subset=["Year"]).copy()
    df_prod["Year"] = df_prod["Year"].astype(int)
    df_prod["Mton"] = pd.to_numeric(df_prod["Mton"], errors="coerce").fillna(0.0)

    df_vol_raw = pd.read_excel(xl, sheet_name="Mm3 year", header=0)
    df_vol_raw.columns = ["Year", "Q_pit", "Q_gb", "Q_dams", "Q_p2proc", "Q_disch", "Q_leak"]
    df_vol = df_vol_raw.dropna(subset=["Year"]).copy()
    df_vol["Year"] = df_vol["Year"].astype(int)

    df_mix_raw = pd.read_excel(xl, sheet_name="ORE MIXES ", header=None)
    mix_rows = []
    for i in range(10, len(df_mix_raw)):
        row = df_mix_raw.iloc[i]
        try:
            mix_rows.append({"period": float(row.iloc[1]),
                             "frac_kir": float(row.iloc[4]),
                             "frac_lev": float(row.iloc[5])})
        except (TypeError, ValueError):
            pass
    df_mix = pd.DataFrame(mix_rows)

    df_sep_raw = pd.read_excel(xl, sheet_name="SEP83 (SP27) ", header=None)
    sep_years = []
    for v in df_sep_raw.iloc[2, 1:].values:
        try:    sep_years.append(int(float(v)))
        except: sep_years.append(None)

    SEP_SPECIES = {
        "SO4": "sulfat", "Ca": "ca [mg", "Cl": "klorid",
        "NO3": "no3-n",  "NH4": "nh4-n",  "Cu": "cu [",
        "Ni":  "ni [",   "Zn": "zn [",    "Co": "co [",
        "Mo":  "mo [",   "As": "as [",    "Cr": "cr [",
    }
    pit_conc = {p: {} for p in SEP_SPECIES}
    for ri in range(3, len(df_sep_raw)):
        label = str(df_sep_raw.iloc[ri, 0]).strip().lower()
        for param, key in SEP_SPECIES.items():
            if key in label:
                for ci, yr in enumerate(sep_years):
                    if yr is not None:
                        try: pit_conc[param][yr] = float(df_sep_raw.iloc[ri, ci + 1])
                        except: pass
    pit_mean = {p: float(np.mean(list(v.values()))) if v else 1.0
                for p, v in pit_conc.items()}

    # Find the SVA79 sheet dynamically (name contains Swedish chars that can encode differently)
    sva_sheet = next((s for s in xl.sheet_names if "SVA79" in s), None)
    if sva_sheet is None:
        raise ValueError(f"Cannot find SVA79 sheet. Available sheets: {xl.sheet_names}")
    df_sva_raw = pd.read_excel(xl, sheet_name=sva_sheet, header=None)
    sva_years = []
    for v in df_sva_raw.iloc[2, 1:].values:
        try:    sva_years.append(int(float(v)))
        except: sva_years.append(None)
    GB_SPECIES = {"SO4": "Sulfat", "Ca": "Ca", "Cl": "Klorid",
                  "NO3": "NO3-N",  "NH4": "NH4-N", "Cu": "Cu",
                  "Ni": "Ni",      "Zn": "Zn",     "Co": "Co",
                  "Mo": "Mo",      "As": "As",      "Cr": "Cr"}
    gb_conc = {p: {} for p in GB_SPECIES}
    for ri in range(3, len(df_sva_raw)):
        label = str(df_sva_raw.iloc[ri, 0]).strip()
        for param, key in GB_SPECIES.items():
            if label.lower().startswith(key.lower()):
                for ci, yr in enumerate(sva_years):
                    if yr is not None:
                        try: gb_conc[param][yr] = float(df_sva_raw.iloc[ri, ci + 1])
                        except: pass
    gb_mean = {p: float(np.mean(list(v.values()))) if v else 0.0
               for p, v in gb_conc.items()}

    print(f"  Monitoring rows : {len(df_mon)}")
    print(f"  Production rows : {len(df_prod)}")
    print(f"  Volume rows     : {len(df_vol)}")
    print(f"  Ore-mix rows    : {len(df_mix)}")
    return df_mon, df_prod, df_vol, df_mix, pit_conc, pit_mean, gb_conc, gb_mean


# ── INPUT BUILDER ─────────────────────────────────────────────

def _get_vol(df_vol, year_int, col, fallback):
    row = df_vol[df_vol["Year"] == year_int]
    return float(row[col].iloc[0]) if len(row) > 0 else fallback

def _get_mix(df_mix, period):
    row = df_mix[np.isclose(df_mix["period"], period, atol=0.01)]
    if len(row) > 0:
        return float(row["frac_kir"].iloc[0]), float(row["frac_lev"].iloc[0])
    yr = int(period)
    ann = df_mix[(df_mix["period"] >= yr) & (df_mix["period"] < yr + 1)]
    if len(ann) > 0:
        return float(ann["frac_kir"].mean()), float(ann["frac_lev"].mean())
    return 0.42, 0.50

def build_inputs(df_prod, df_vol, df_mix, half_years):
    avg_pit   = df_vol["Q_pit"].mean()
    avg_gb    = df_vol["Q_gb"].mean()
    avg_dams  = df_vol["Q_dams"].mean()
    avg_disch = df_vol["Q_disch"].mean()
    avg_leak  = df_vol["Q_leak"].mean()
    rows = []
    for yh in half_years:
        yr   = int(yh)
        half = 1 if (yh - yr) > 0.1 else 0
        ann_prod = df_prod.loc[df_prod["Year"] == yr, "Mton"]
        ann_mton = float(ann_prod.iloc[0]) if len(ann_prod) > 0 else 4.267
        prod_half_ton = ann_mton * (7/12 if half else 5/12) * 1e6
        frac_kir, frac_lev = _get_mix(df_mix, yh)
        rows.append({
            "period": yh, "half": half, "prod_ton": prod_half_ton,
            "frac_kir": frac_kir, "frac_lev": frac_lev,
            "Q_pit":   _get_vol(df_vol, yr, "Q_pit",   avg_pit)   / 2,
            "Q_gb":    _get_vol(df_vol, yr, "Q_gb",    avg_gb)    / 2,
            "Q_dams":  _get_vol(df_vol, yr, "Q_dams",  avg_dams)  / 2,
            "Q_disch": _get_vol(df_vol, yr, "Q_disch", avg_disch) / 2,
            "Q_leak":  _get_vol(df_vol, yr, "Q_leak",  avg_leak)  / 2,
        })
    return pd.DataFrame(rows).sort_values("period").reset_index(drop=True)


# ── MASS-BALANCE RECURRENCE ───────────────────────────────────

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


# ── PLOT ──────────────────────────────────────────────────────

def plot_results(results, df_mon):
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "font.size": 9.5,
        "figure.facecolor": "#0f1117", "axes.facecolor": "#1a1d27",
        "axes.edgecolor": "#444", "axes.labelcolor": "#e0e0e0",
        "xtick.color": "#aaa", "ytick.color": "#aaa",
        "text.color": "#e0e0e0", "grid.color": "#2a2d3a",
        "grid.linewidth": 0.6, "axes.grid": True,
        "axes.spines.top": False, "axes.spines.right": False,
    })
    params = list(results.keys())
    n_cols = 4
    n_rows = (len(params) + n_cols - 1) // n_cols
    fig, axes = plt.subplots(n_rows, n_cols, figsize=(24, n_rows * 5.5), squeeze=False)
    fig.suptitle(
        "Leveäniemi Mine — Method 1: Mass-Balance Model  |  Validation 2016–2025  &  Forecast 2026–2030\n"
        "White line = SVA79 actual monitoring   |   Coloured line = mass-balance prediction   |   Dashed = forecast",
        fontsize=12, fontweight="bold", y=0.995, color="#f0f0f0",
    )
    for idx, param in enumerate(params):
        ax     = axes[idx // n_cols][idx % n_cols]
        colour = COLOURS[param]
        unit   = UNITS[param]
        dd_col = DD_COLS[param]
        if dd_col in df_mon.columns:
            mon_vals = pd.to_numeric(df_mon[dd_col], errors="coerce")
            valid    = mon_vals.notna()
            ax.plot(df_mon.loc[valid, "Period"], mon_vals[valid],
                    color="white", lw=1.8, zorder=6, label="SVA79 monitoring",
                    marker="o", markersize=5,
                    markerfacecolor="white", markeredgecolor="#555", markeredgewidth=0.7)
        val = results[param]["validation"]
        ax.plot(val["period"], val["pred"], color=colour, lw=2.0, zorder=5,
                label="Mass-balance (LK)")
        fct = results[param]["forecast"]
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
    plt.savefig(FIGURE_OUT, dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
    if IN_COLAB:
        plt.show()      # displays inline in Colab
    else:
        plt.close()
    print(f"  Figure saved -> {FIGURE_OUT}")


# ── EXCEL EXPORT ──────────────────────────────────────────────

def export_excel(results, df_mon):
    with pd.ExcelWriter(EXCEL_OUT, engine="xlsxwriter") as writer:
        wb  = writer.book
        hdr = wb.add_format({"bold": True, "bg_color": "#1f4e79", "font_color": "white", "border": 1})
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
                per     = row["period"]
                mon_row = df_mon[np.isclose(df_mon["Period"], per, atol=0.01)]
                act     = None
                if len(mon_row) > 0 and dd_col in df_mon.columns:
                    a = pd.to_numeric(mon_row[dd_col].iloc[0], errors="coerce")
                    act = float(a) if pd.notna(a) else None
                pred = round(float(row["pred"]), 4)
                diff = round(pred - act, 4) if act is not None else ""
                if act is not None:
                    errors.append(pred - act)
                rows.append({
                    "Period": per,
                    "Season": "Summer" if int(row["half"]) == 1 else "Winter",
                    f"Actual ({unit})":  round(act, 4) if act is not None else "",
                    f"Method1 ({unit})": pred,
                    "Difference": diff,
                })
            for _, row in res["forecast"].iterrows():
                rows.append({
                    "Period": row["period"],
                    "Season": "Summer" if int(row["half"]) == 1 else "Winter",
                    f"Actual ({unit})":  "—",
                    f"Method1 ({unit})": round(float(row["pred"]), 4),
                    "Difference": "—",
                })
            df_out = pd.DataFrame(rows)
            df_out.to_excel(writer, sheet_name=param[:31], index=False, startrow=2)
            ws = writer.sheets[param[:31]]
            ws.write(0, 0, f"{param} — Mass-Balance LK", ttl)
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
                "Kir rate (g/ton)": KIR_RATES_GT[param],
                "MAE": round(mae, 4) if not np.isnan(mae) else "",
                "RMSE": round(rmse, 4) if not np.isnan(rmse) else "",
                "n_valid": len(errors),
            })
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False, startrow=2)
        ws = writer.sheets["Summary"]
        ws.write(0, 0, "Mass-Balance Model — Validation Summary", ttl)
        for ci, col in enumerate(pd.DataFrame(summary_rows).columns):
            ws.write(2, ci, col, hdr)
            ws.set_column(ci, ci, 22)
    print(f"  Excel saved -> {EXCEL_OUT}")


# ── MAIN ──────────────────────────────────────────────────────

def main(params_file=None):
    # In Colab, PARAMS_FILE is set by the upload cell above.
    # If running locally, pass the filename directly: main("parameters_used2.xlsx")
    if params_file is None:
        try:
            params_file = PARAMS_FILE   # set by Colab upload cell
        except NameError:
            params_file = "parameters_used2.xlsx"

    print("=" * 60)
    print("Leveäniemi — Method 1: Mass-Balance Model")
    print("=" * 60)
    print("\nLoading data ...")
    df_mon, df_prod, df_vol, df_mix, pit_conc, pit_mean, gb_conc, gb_mean = load_data(params_file)

    val_periods = sorted(df_mon["Period"].dropna().unique().tolist())
    print(f"\nValidation : {val_periods[0]} → {val_periods[-1]}")
    print(f"Forecast   : {FORECAST_HALF_YEARS[0]} → {FORECAST_HALF_YEARS[-1]}")

    df_val_inp = build_inputs(df_prod, df_vol, df_mix, val_periods)
    df_fct_inp = build_inputs(df_prod, df_vol, df_mix, FORECAST_HALF_YEARS)

    results = {}
    for param in list(LEV_RATES_GT.keys()):
        dd_col = DD_COLS[param]
        unit   = UNITS[param]
        if dd_col in df_mon.columns:
            first = pd.to_numeric(df_mon[df_mon["Period"] == val_periods[0]][dd_col], errors="coerce")
            seed  = float(first.iloc[0]) if (len(first) > 0 and pd.notna(first.iloc[0])) else 1.0
        else:
            seed = 1.0
        val_preds = run_mass_balance(df_val_inp, param, seed, pit_conc, pit_mean, gb_mean)
        fct_preds = run_mass_balance(df_fct_inp, param, val_preds[-1], pit_conc, pit_mean, gb_mean)
        df_val_out = df_val_inp.copy(); df_val_out["pred"] = val_preds
        df_fct_out = df_fct_inp.copy(); df_fct_out["pred"] = fct_preds
        errors = []
        for per, pr in zip(df_val_out["period"], val_preds):
            if dd_col in df_mon.columns:
                mon_r = df_mon[np.isclose(df_mon["Period"], per, atol=0.01)]
                if len(mon_r) > 0:
                    a = pd.to_numeric(mon_r[dd_col].iloc[0], errors="coerce")
                    if pd.notna(a):
                        errors.append(pr - float(a))
        mae  = np.mean(np.abs(errors)) if errors else np.nan
        rmse = np.sqrt(np.mean(np.array(errors)**2)) if errors else np.nan
        print(f"  {param:4s}  seed={seed:.3f}  range=[{val_preds.min():.3f}, {val_preds.max():.3f}]"
              f"  MAE={mae:.3f}  RMSE={rmse:.3f}  [{unit}]")
        results[param] = {"validation": df_val_out, "forecast": df_fct_out}

    print("\nGenerating figure ...")
    plot_results(results, df_mon)

    print("Exporting Excel ...")
    export_excel(results, df_mon)

    # Offer downloads when running in Colab
    if IN_COLAB and colab_files is not None:
        colab_files.download(FIGURE_OUT)
        colab_files.download(EXCEL_OUT)

    print("\n" + "=" * 60)
    print(f"DONE  >>  {FIGURE_OUT}  |  {EXCEL_OUT}")
    print("=" * 60)
    return results


# Run!
results = main()
