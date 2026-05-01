
# =============================================================================
# CELL 1 — Install packages (run this first, then restart runtime if prompted)
# =============================================================================
# !pip install pandas numpy scikit-learn matplotlib openpyxl xlsxwriter


# =============================================================================
# CELL 2 — Upload the Excel data file
# =============================================================================
# from google.colab import files
# uploaded = files.upload()
# # Select: Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx
# print("Uploaded:", list(uploaded.keys()))


# =============================================================================
# CELL 3 — Imports
# =============================================================================
import warnings; warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import TimeSeriesSplit
from sklearn.metrics import r2_score
from sklearn.pipeline import Pipeline
print("All imports OK")


# =============================================================================
# CELL 4 — Configuration: file name, block locations, column indices
# =============================================================================

# ---- UPDATE THIS if your file uploaded with a different name ----
EXCEL_FILE = "Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx"
SHEET_NAME = "Process water"

# Where each contaminant block lives in the sheet (0-based row indices)
BLOCKS = {
    "Cu":  dict(label="Cu",  unit="ug/L",  first_row=42,  last_row=74),
    "NH4": dict(label="NH4", unit="mg/L",  first_row=81,  last_row=113),
    "Cl":  dict(label="Cl",  unit="mg/L",  first_row=121, last_row=153),
    "Ni":  dict(label="Ni",  unit="ug/L",  first_row=160, last_row=192),
}

# Column indices (0-based), same for all blocks
COL_YEAR      = 0
COL_HALF_YEAR = 1   # 1=summer (processing season), 0=winter
COL_PROD_GM   = 3   # Ore production volume  GM (Mton)
COL_LEACH_GM  = 4   # Leaching load from GM ore
COL_PROD_GK   = 5   # Ore production volume GK (Mton)
COL_LEACH_GK  = 6   # Leaching load from GK ore
COL_LEACH_GL  = 7   # Leaching load from GL ore
COL_PIT_PUMP  = 8   # Leveaniemi pit pump volume (Mm3)
COL_MODELLED  = 9   # Consultant's modelled concentration  <-- TARGET
COL_STORAGE   = 31  # Total storage concentration  <-- RECURRENCE STATE

TRAIN_UNTIL   = 2025.5   # Last training row (inclusive)
FORECAST_FROM = 2026.0   # First forecast row

N_MC_RUNS  = 300    # Monte Carlo iterations
MC_PERTURB = 0.15   # +/-15% perturbation on leaching rates

RF_PARAMS = dict(n_estimators=500, random_state=42, n_jobs=-1,
                 min_samples_leaf=2, max_features="sqrt")
N_CV_SPLITS = 4

FIGURE_OUT = "forecast_figure.png"
EXCEL_OUT  = "forecast_results.xlsx"

# LK (Leveaniemi-Kiruna) leaching rates — 50/50 average of Lev + Kiruna rates
# Source: ore-tail leach calc sheet in the Excel workbook
LK_LEACH_RATES = {
    "Cu":  (1.358 + 11.668) / 2,    # 6.513 kg/Mton
    "NH4": (0.0   + 3047.3) / 2,    # 1523.65 kg/Mton (Lev = n/a)
    "Cl":  (154444 + 285406) / 2,   # 219925 kg/Mton
    "Ni":  (0.631 + 1.174)  / 2,    # 0.9025 kg/Mton
}
LK_PROD_WINTER = 1.6667   # Mton (placeholder — update when confirmed)
LK_PROD_SUMMER = 2.5667   # Mton (placeholder — update when confirmed)

COLOURS = {"Cu": "#00c8e8", "NH4": "#f77f00", "Cl": "#7bc67e", "Ni": "#c77dff"}

FEATURE_COLS = [
    "half_year", "prod_gm", "leach_gm",
    "prod_gk",  "leach_gk", "leach_gl",
    "prod_lk",  "leach_lk",
    "pit_pump", "prev_storage",
]
print("Configuration ready.")


# =============================================================================
# CELL 5 — Helper functions: parse_block, inject_lk_forecast, build_features
# =============================================================================

def _safe_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return np.nan


def parse_block(df_raw, param):
    """
    Read one contaminant block from the raw Excel sheet.
    Returns a clean DataFrame with all input columns and the target column.

    prod_lk and leach_lk are set to 0 because the consultant never modelled
    the LK (Leveaniemi-Kiruna) ore combination. They are filled in later for
    the forecast rows only.
    """
    cfg  = BLOCKS[param]
    rows = df_raw.iloc[cfg["first_row"] : cfg["last_row"] + 1].copy()
    data = []
    for _, row in rows.iterrows():
        year = _safe_float(row.iloc[COL_YEAR])
        if np.isnan(year) or not (2010 <= year <= 2035):
            continue
        hy = row.iloc[COL_HALF_YEAR]
        hy = 0 if pd.isna(hy) else int(hy)
        data.append(dict(
            year          = year,
            half_year     = hy,
            prod_gm       = _safe_float(row.iloc[COL_PROD_GM]),
            leach_gm      = _safe_float(row.iloc[COL_LEACH_GM]),
            prod_gk       = _safe_float(row.iloc[COL_PROD_GK]),
            leach_gk      = _safe_float(row.iloc[COL_LEACH_GK]),
            leach_gl      = _safe_float(row.iloc[COL_LEACH_GL]),
            prod_lk       = 0.0,   # LK not in consultant model
            leach_lk      = 0.0,   # LK not in consultant model
            pit_pump      = _safe_float(row.iloc[COL_PIT_PUMP]),
            modelled_conc = _safe_float(row.iloc[COL_MODELLED]),
            storage_conc  = _safe_float(row.iloc[COL_STORAGE]),
        ))
    return pd.DataFrame(data).sort_values("year").reset_index(drop=True)


def inject_lk_forecast(df_forecast, param):
    """
    Fill in LK production and leaching for the 2026-2030 forecast rows.
    Uses LK_PROD_* and LK_LEACH_RATES from configuration above.
    When Agnes confirms actual LK volumes, update LK_PROD_WINTER/SUMMER.
    """
    df = df_forecast.copy()
    rate = LK_LEACH_RATES[param]
    df["prod_lk"]  = df["half_year"].apply(
        lambda h: LK_PROD_SUMMER if h == 1 else LK_PROD_WINTER)
    df["leach_lk"] = df["prod_lk"] * rate
    return df


def build_features(df):
    """
    Add prev_storage (lag-1 of storage_conc).
    The first row has no previous row, so it is dropped.
    This lag is the KEY recurrence state that mimics the consultant's formula.
    """
    df = df.copy()
    df["prev_storage"] = df["storage_conc"].shift(1)
    df = df.dropna(subset=["prev_storage", "modelled_conc"]).reset_index(drop=True)
    return df

print("Helper functions defined.")


# =============================================================================
# CELL 6 — Model functions: train_model, run_forecast, monte_carlo
# =============================================================================

def train_model(df_train, param):
    """
    Train a Random Forest regressor on the historical data (2015-2025).

    Uses temporal cross-validation (TimeSeriesSplit) so that the model
    is always tested on data AFTER its training data — never the reverse.

    Returns the trained pipeline, CV R2 scores, and in-sample R2.
    """
    X = df_train[FEATURE_COLS].values
    y = df_train["modelled_conc"].values

    pipe = Pipeline([
        ("scaler", StandardScaler()),          # normalize all features to same scale
        ("rf",     RandomForestRegressor(**RF_PARAMS)),
    ])

    tscv = TimeSeriesSplit(n_splits=N_CV_SPLITS)
    cv_scores = []
    for tr_idx, val_idx in tscv.split(X):
        pipe.fit(X[tr_idx], y[tr_idx])
        cv_scores.append(r2_score(y[val_idx], pipe.predict(X[val_idx])))

    pipe.fit(X, y)   # final fit on ALL training data
    in_sample_r2 = r2_score(y, pipe.predict(X))

    cv_mean = float(np.mean(cv_scores))
    print(f"  [{param}]  CV R2 mean: {cv_mean:.3f}  |  "
          + "  ".join(f"{s:.3f}" for s in cv_scores)
          + f"  |  in-sample R2: {in_sample_r2:.3f}")
    if cv_mean < 0.5:
        print(f"  WARNING [{param}]: CV R2 < 0.5")
    return pipe, cv_scores, in_sample_r2


def run_forecast(pipe, df_forecast, seed_storage, leach_scale=1.0):
    """
    Recurrently predict concentrations for 2026-2030.

    Key idea: each prediction is fed back as 'prev_storage' for the next step.
    This mirrors the consultant's recurrence formula exactly.

    leach_scale: used by Monte Carlo to perturb leaching rates +/- 15%.
    """
    predictions  = []
    prev_storage = seed_storage
    for _, row in df_forecast.iterrows():
        x = np.array([[
            row["half_year"],
            row["prod_gm"],
            row["leach_gm"]  * leach_scale,
            row["prod_gk"],
            row["leach_gk"]  * leach_scale,
            row["leach_gl"]  * leach_scale,
            row["prod_lk"],
            row["leach_lk"]  * leach_scale,
            row["pit_pump"],
            prev_storage,
        ]])
        pred         = pipe.predict(x)[0]
        predictions.append(pred)
        prev_storage = pred
    return np.array(predictions)


def monte_carlo(pipe, df_forecast, seed_storage, rng):
    """
    Run N_MC_RUNS forecasts, each with a randomly perturbed leaching rate.
    Perturbation is uniform +/- MC_PERTURB (15%) applied to ALL leaching columns.
    Returns P10, P50, P90 across all runs.
    """
    all_runs = np.zeros((N_MC_RUNS, len(df_forecast)))
    for i in range(N_MC_RUNS):
        scale = rng.uniform(1.0 - MC_PERTURB, 1.0 + MC_PERTURB)
        all_runs[i] = run_forecast(pipe, df_forecast, seed_storage, leach_scale=scale)
    return dict(
        all_runs = all_runs,
        p10      = np.percentile(all_runs, 10, axis=0),
        p50      = np.percentile(all_runs, 50, axis=0),
        p90      = np.percentile(all_runs, 90, axis=0),
    )

print("Model functions defined.")


# =============================================================================
# CELL 7 — Load the Excel file and parse all four contaminant blocks
# =============================================================================

print("Loading Excel workbook (this takes ~20 seconds) ...")
df_raw = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME,
                       header=None, engine="openpyxl")
print(f"Sheet loaded: {df_raw.shape[0]} rows x {df_raw.shape[1]} columns")

parsed = {}
for param in BLOCKS:
    df = parse_block(df_raw, param)
    df = build_features(df)
    parsed[param] = df
    print(f"  {param}: {len(df)} rows  "
          f"({df['year'].min()} to {df['year'].max()})")

print("\nAll blocks parsed successfully.")


# =============================================================================
# CELL 8 — Train one Random Forest model per contaminant
# =============================================================================

rng     = np.random.default_rng(seed=42)
trained = {}

print("Training models ...\n")
for param in BLOCKS:
    df       = parsed[param]
    df_train = df[df["year"] <= TRAIN_UNTIL].copy()
    pipe, cv_scores, in_sample_r2 = train_model(df_train, param)
    trained[param] = dict(
        pipe         = pipe,
        cv_scores    = cv_scores,
        in_sample_r2 = in_sample_r2,
        df_train     = df_train,
    )

print("\nAll models trained.")


# =============================================================================
# CELL 9 — Run Monte Carlo forecasts for 2026-2030 (incl. LK)
# =============================================================================

print("Running Monte Carlo simulations ...\n")
results = {}

for param in BLOCKS:
    df           = parsed[param]
    df_forecast  = df[df["year"] >= FORECAST_FROM].copy()
    df_forecast  = inject_lk_forecast(df_forecast, param)   # add LK loads
    seed_storage = trained[param]["df_train"]["storage_conc"].iloc[-1]

    lk_rate = LK_LEACH_RATES[param]
    print(f"  {param}: LK leach rate = {lk_rate:.4f} kg/Mton  "
          f"(winter prod={LK_PROD_WINTER}, summer prod={LK_PROD_SUMMER} Mton -- PLACEHOLDER)")

    mc = monte_carlo(trained[param]["pipe"], df_forecast, seed_storage, rng)
    print(f"  {param}: P50 range {mc['p50'].min():.4f} -- {mc['p50'].max():.4f}"
          f"  {BLOCKS[param]['unit']}\n")

    results[param] = dict(
        df_train     = trained[param]["df_train"],
        df_forecast  = df_forecast,
        cv_scores    = trained[param]["cv_scores"],
        in_sample_r2 = trained[param]["in_sample_r2"],
        mc           = mc,
    )

print("All Monte Carlo runs complete.")


# =============================================================================
# CELL 10 — Generate the 4-panel forecast figure
# =============================================================================

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 11, "axes.titlesize": 13,
    "axes.spines.top": False, "axes.spines.right": False,
    "figure.facecolor": "#0f1117", "axes.facecolor": "#1a1d27",
    "axes.edgecolor": "#444", "axes.labelcolor": "#e0e0e0",
    "xtick.color": "#aaa", "ytick.color": "#aaa", "text.color": "#e0e0e0",
    "grid.color": "#2a2d3a", "grid.linewidth": 0.6, "axes.grid": True,
    "legend.framealpha": 0.2, "legend.facecolor": "#1a1d27",
})

params = list(results.keys())
fig    = plt.figure(figsize=(16, 11))
fig.suptitle(
    "Leveaniemi Mine - Process Water Quality Forecast 2026-2030\n"
    "Random Forest with Monte Carlo Uncertainty (P10/P90, n=300, +/-15%)",
    fontsize=15, fontweight="bold", y=0.98, color="#f0f0f0",
)
gs = GridSpec(2, 2, figure=fig, hspace=0.42, wspace=0.32,
              left=0.07, right=0.97, top=0.91, bottom=0.09)
pos_map = {"Cu": gs[0,0], "NH4": gs[0,1], "Cl": gs[1,0], "Ni": gs[1,1]}

for param in params:
    ax      = fig.add_subplot(pos_map[param])
    res     = results[param]
    colour  = COLOURS[param]
    hist    = res["df_train"]
    fcst_df = res["df_forecast"]
    mc      = res["mc"]
    yrs     = fcst_df["year"].values

    ax.plot(hist["year"], hist["modelled_conc"],
            color=colour, lw=2.0, alpha=0.9, label="Consultant model (hist.)")
    ax.plot(fcst_df["year"], fcst_df["modelled_conc"],
            color="#bbbbbb", lw=1.4, ls="--", alpha=0.7, label="Consultant extrap.")
    ax.fill_between(yrs, mc["p10"], mc["p90"],
                    color=colour, alpha=0.22, label="RF P10-P90")
    ax.plot(yrs, mc["p50"],
            color=colour, lw=2.5, label="RF P50 forecast")
    ax.axvline(x=FORECAST_FROM - 0.5, color="#666", lw=0.9, ls=":")
    ax.text(FORECAST_FROM - 0.4, ax.get_ylim()[1], "Forecast >>",
            fontsize=8, color="#888", ha="left", va="top")

    unit   = BLOCKS[param]["unit"]
    cv_r2  = float(np.mean(res["cv_scores"]))
    badge  = "#44aa66" if cv_r2 >= 0.7 else ("#f0a030" if cv_r2 >= 0.5 else "#ee4444")
    ax.set_title(f"{BLOCKS[param]['label']}  [{unit}]", color="#f0f0f0", pad=6)
    ax.set_xlabel("Year", labelpad=4)
    ax.set_ylabel(f"Concentration  ({unit})", labelpad=4)
    ax.set_xlim(2014, 2031)
    ax.text(0.02, 0.97, f"CV R2 = {cv_r2:.3f}", transform=ax.transAxes,
            fontsize=9.5, fontweight="bold", va="top", ha="left", color=badge,
            bbox=dict(boxstyle="round,pad=0.35", fc="#22253a", ec="none", alpha=0.85))
    if param == params[0]:
        ax.legend(loc="upper right", fontsize=8.5)

handles = [
    plt.Line2D([0],[0], color=list(COLOURS.values())[0], lw=2.0,
               label="Consultant model (hist.)"),
    plt.Line2D([0],[0], color="#bbbbbb", lw=1.4, ls="--",
               label="Consultant extrapolation (2026-2030)"),
    mpatches.Patch(alpha=0.35, color="#888888",
                   label="RF P10-P90  (MC +/-15%, n=300)"),
    plt.Line2D([0],[0], color="#ffffff", lw=2.5, label="RF P50 forecast"),
]
fig.legend(handles=handles, loc="lower center", ncol=4, fontsize=9.5,
           framealpha=0.2, facecolor="#1a1d27", edgecolor="#444",
           bbox_to_anchor=(0.5, 0.01))

plt.savefig(FIGURE_OUT, dpi=180, bbox_inches="tight",
            facecolor=fig.get_facecolor())
plt.show()
print(f"Figure saved: {FIGURE_OUT}")


# =============================================================================
# CELL 11 — Export results to Excel
# =============================================================================

cv_summary = []
for param, res in results.items():
    cv_summary.append({
        "Parameter":                param,
        "Unit":                     BLOCKS[param]["unit"],
        "Training rows":            len(res["df_train"]),
        "CV R2 mean":               round(float(np.mean(res["cv_scores"])), 4),
        "CV R2 std":                round(float(np.std(res["cv_scores"])),  4),
        "In-sample R2":             round(res["in_sample_r2"], 4),
        "MC runs":                  N_MC_RUNS,
        "MC perturbation %":        int(MC_PERTURB * 100),
        "LK leach rate (kg/Mton)":  round(LK_LEACH_RATES[param], 4),
        "LK prod winter (Mton)":    LK_PROD_WINTER,
        "LK prod summer (Mton)":    LK_PROD_SUMMER,
        "LK assumption":            "50/50 Leveaniemi+Kiruna (PLACEHOLDER)",
    })

with pd.ExcelWriter(EXCEL_OUT, engine="xlsxwriter") as writer:
    wb  = writer.book
    hdr = wb.add_format({"bold": True, "bg_color": "#1f4e79",
                         "font_color": "white", "border": 1, "align": "center"})
    num = wb.add_format({"num_format": "0.0000", "border": 1})
    txt = wb.add_format({"border": 1})
    ttl = wb.add_format({"bold": True, "font_size": 12, "font_color": "#1f4e79"})

    for param, res in results.items():
        mc      = res["mc"]
        fcst_df = res["df_forecast"]
        unit    = BLOCKS[param]["unit"]
        rows_out = []
        for i, (_, row) in enumerate(fcst_df.iterrows()):
            rows_out.append({
                "Year":                  row["year"],
                "Season":                "Summer" if int(row["half_year"]) == 1 else "Winter",
                f"P10 ({unit})":         mc["p10"][i],
                f"P50 ({unit})":         mc["p50"][i],
                f"P90 ({unit})":         mc["p90"][i],
                f"Consultant ({unit})":  row["modelled_conc"],
            })
        df_out = pd.DataFrame(rows_out)
        df_out.to_excel(writer, sheet_name=param, index=False, startrow=2)
        ws = writer.sheets[param]
        ws.write(0, 0, f"{param} -- ML Forecast 2026-2030 | RF + MC (n={N_MC_RUNS})", ttl)
        for ci, col_name in enumerate(df_out.columns):
            ws.write(2, ci, col_name, hdr)
            ws.set_column(ci, ci, max(16, len(col_name) + 3))
        for ri in range(len(df_out)):
            for ci in range(len(df_out.columns)):
                v = df_out.iloc[ri, ci]
                ws.write(ri + 3, ci, v, num if isinstance(v, float) else txt)

    pd.DataFrame(cv_summary).to_excel(writer, sheet_name="CV_Diagnostics",
                                      index=False, startrow=2)
    ws = writer.sheets["CV_Diagnostics"]
    ws.write(0, 0, "Cross-Validation Diagnostics -- Random Forest Models", ttl)

print(f"Excel saved: {EXCEL_OUT}")


# =============================================================================
# CELL 12 — Download output files (Colab only)
# =============================================================================
# from google.colab import files
# files.download(FIGURE_OUT)
# files.download(EXCEL_OUT)
# print("Downloads started.")
