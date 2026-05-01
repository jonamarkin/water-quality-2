"""
water_quality_model.py
======================
Leveaniemi Mine -- Process Water Quality ML Forecast
Master's Thesis, LKAB / Swedish University

Workflow:
  1. parse_block()    -- extract one contaminant block from the Excel sheet
  2. build_features() -- feature matrix with lagged recurrence state
  3. train_model()    -- Random Forest + temporal cross-validation
  4. run_forecast()   -- recurrent prediction for 2026-2030
  5. monte_carlo()    -- 300-run MC with +-15% leaching perturbation
  6. plot_results()   -- thesis-quality 4-panel matplotlib figure
  7. export_excel()   -- formatted Excel table of predictions

Usage:   python water_quality_model.py
Outputs: forecast_figure.png   forecast_results.xlsx
"""

import sys, io, platform

# Force UTF-8 only on Windows (cp1252 terminals) -- not needed on Colab/Linux
if platform.system() == "Windows":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")        # headless / no display needed
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import TimeSeriesSplit
from sklearn.metrics import r2_score
from sklearn.pipeline import Pipeline

# ===========================================================================
# CONFIGURATION
# ===========================================================================

EXCEL_FILE = "Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx"
SHEET_NAME = "Process water"

# Block definitions: first_row/last_row are 0-based row indices in the sheet
BLOCKS = {
    "Cu":  dict(label="Cu",  unit="ug/L",  first_row=42,  last_row=74),
    "NH4": dict(label="NH4", unit="mg/L",  first_row=81,  last_row=113),
    "Cl":  dict(label="Cl",  unit="mg/L",  first_row=121, last_row=153),
    "Ni":  dict(label="Ni",  unit="ug/L",  first_row=160, last_row=192),
}

# Column indices (0-based), identical for all four blocks
COL_YEAR      = 0
COL_HALF_YEAR = 1   # 1=summer processing season, 0=winter; NaN treated as 0
COL_PROD_GM   = 3   # Production Mton GM
COL_LEACH_GM  = 4   # Process leaching GM
COL_PROD_GK   = 5   # Production Mton GK
COL_LEACH_GK  = 6   # Process leaching GK
COL_LEACH_GL  = 7   # Process leaching GL
COL_PIT_PUMP  = 8   # Leveaniemi pit pump volume (Mm3)
COL_MODELLED  = 9   # Modelled concentration  <-- TARGET
COL_STORAGE   = 31  # Total storage concentration  <-- RECURRENCE STATE

TRAIN_UNTIL   = 2025.5   # inclusive: last training row
FORECAST_FROM = 2026.0   # first forecast row

N_MC_RUNS   = 300    # Monte Carlo iterations
MC_PERTURB  = 0.15   # +/-15% uniform perturbation on leaching rates

RF_PARAMS = dict(
    n_estimators=500,
    random_state=42,
    n_jobs=-1,
    min_samples_leaf=2,
    max_features="sqrt",
)
N_CV_SPLITS = 4   # TimeSeriesSplit folds

FIGURE_OUT = "forecast_figure.png"
EXCEL_OUT  = "forecast_results.xlsx"

# ---------------------------------------------------------------------------
# LK (Leveaniemi-Kiruna) ore combination
# ---------------------------------------------------------------------------
# The consultant never modelled LK. Leaching rates are derived from the
# ore-tail leach calc sheet as a 50/50 average of Leveaniemi + Kiruna rates.
# For NH4: Leveaniemi rate is 'na' (not measured), so only Kiruna contributes
#          -> LK_NH4 = 0.5 * 0 + 0.5 * 3047.3 = 1523.65 kg/Mton
# Production volumes use GK values as placeholder (update when known).
LK_LEACH_RATES = {
    "Cu":  (1.358 + 11.668) / 2,          # 6.513  kg/Mton
    "NH4": (0.0   + 3047.3) / 2,          # 1523.65 kg/Mton (Lev = n/a -> 0)
    "Cl":  (154444 + 285406) / 2,         # 219925  kg/Mton
    "Ni":  (0.631 + 1.174)  / 2,          # 0.9025 kg/Mton
}
# Placeholder production volumes (Mton) -- same as GK split until confirmed
LK_PROD_WINTER = 1.6667   # x.0 rows (5-month winter, no processing)
LK_PROD_SUMMER = 2.5667   # x.5 rows (7-month summer, processing season)

# Colour palette (vibrant, distinct hues)
COLOURS = {
    "Cu":  "#00c8e8",
    "NH4": "#f77f00",
    "Cl":  "#7bc67e",
    "Ni":  "#c77dff",
}

# Feature columns -- includes LK features (zero during training)
FEATURE_COLS = [
    "half_year", "prod_gm", "leach_gm",
    "prod_gk",  "leach_gk", "leach_gl",
    "prod_lk",  "leach_lk",
    "pit_pump", "prev_storage",
]


# ===========================================================================
# 1. PARSE BLOCK
# ===========================================================================

def _safe_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return np.nan


def parse_block(df_raw: pd.DataFrame, param: str) -> pd.DataFrame:
    """
    Extract one contaminant block from the raw sheet.
    Returns a tidy DataFrame with columns:
      year, half_year, prod_gm, leach_gm, prod_gk, leach_gk, leach_gl,
      prod_lk, leach_lk, pit_pump, modelled_conc, storage_conc

    NOTE: prod_lk and leach_lk are set to 0 for ALL consultant rows because
    the consultant never modelled the LK (Leveaniemi-Kiruna) ore combination.
    They become non-zero only in the forecast period via inject_lk_forecast().
    """
    cfg  = BLOCKS[param]
    rows = df_raw.iloc[cfg["first_row"] : cfg["last_row"] + 1].copy()

    data = []
    for _, row in rows.iterrows():
        year = _safe_float(row.iloc[COL_YEAR])
        if np.isnan(year) or not (2010 <= year <= 2035):
            continue

        half_year = row.iloc[COL_HALF_YEAR]
        half_year = 0 if pd.isna(half_year) else int(half_year)

        data.append(dict(
            year         = year,
            half_year    = half_year,
            prod_gm      = _safe_float(row.iloc[COL_PROD_GM]),
            leach_gm     = _safe_float(row.iloc[COL_LEACH_GM]),
            prod_gk      = _safe_float(row.iloc[COL_PROD_GK]),
            leach_gk     = _safe_float(row.iloc[COL_LEACH_GK]),
            leach_gl     = _safe_float(row.iloc[COL_LEACH_GL]),
            prod_lk      = 0.0,    # LK not in consultant model -> zero
            leach_lk     = 0.0,    # LK not in consultant model -> zero
            pit_pump     = _safe_float(row.iloc[COL_PIT_PUMP]),
            modelled_conc= _safe_float(row.iloc[COL_MODELLED]),
            storage_conc = _safe_float(row.iloc[COL_STORAGE]),
        ))

    df = pd.DataFrame(data).sort_values("year").reset_index(drop=True)
    return df


def inject_lk_forecast(df_forecast: pd.DataFrame, param: str) -> pd.DataFrame:
    """
    Populate prod_lk and leach_lk for the 2026-2030 forecast rows.

    Uses:
      - LK_PROD_SUMMER / LK_PROD_WINTER as production volumes (placeholder)
      - LK_LEACH_RATES[param] as the per-tonne leaching rate (50/50 Lev+Kir)

    When Agnes confirms actual LK production volumes, update LK_PROD_* above.
    """
    df = df_forecast.copy()
    lk_rate = LK_LEACH_RATES[param]
    df["prod_lk"]  = df["half_year"].apply(
        lambda h: LK_PROD_SUMMER if h == 1 else LK_PROD_WINTER
    )
    df["leach_lk"] = df["prod_lk"] * lk_rate
    return df


# ===========================================================================
# 2. BUILD FEATURES
# ===========================================================================

def build_features(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add prev_storage (lag-1 of storage_conc). First row (no lag) is dropped.
    """
    df = df.copy()
    df["prev_storage"] = df["storage_conc"].shift(1)
    df = df.dropna(subset=["prev_storage", "modelled_conc"]).reset_index(drop=True)
    return df


# ===========================================================================
# 3. TRAIN MODEL
# ===========================================================================

def train_model(df_train: pd.DataFrame, param: str):
    """
    Train RF pipeline with temporal CV.
    Returns (pipeline, cv_r2_list, in_sample_r2).
    """
    X = df_train[FEATURE_COLS].values
    y = df_train["modelled_conc"].values

    pipe = Pipeline([
        ("scaler", StandardScaler()),
        ("rf",     RandomForestRegressor(**RF_PARAMS)),
    ])

    tscv      = TimeSeriesSplit(n_splits=N_CV_SPLITS)
    cv_scores = []
    for tr_idx, val_idx in tscv.split(X):
        pipe.fit(X[tr_idx], y[tr_idx])
        cv_scores.append(r2_score(y[val_idx], pipe.predict(X[val_idx])))

    # Final fit on ALL training data
    pipe.fit(X, y)
    in_sample_r2 = r2_score(y, pipe.predict(X))

    cv_mean = float(np.mean(cv_scores))
    fold_str = "  ".join(f"{s:.3f}" for s in cv_scores)
    print(f"  [{param}]  CV R2 (mean): {cv_mean:.3f}  |  folds: {fold_str}"
          f"  |  in-sample R2: {in_sample_r2:.3f}")
    if cv_mean < 0.5:
        print(f"  WARNING [{param}]: CV R2 < 0.5 -- consider split-season sub-models")

    return pipe, cv_scores, in_sample_r2


# ===========================================================================
# 4. RUN FORECAST  (single deterministic run)
# ===========================================================================

def run_forecast(pipe, df_forecast: pd.DataFrame,
                 seed_storage: float,
                 leach_scale: float = 1.0) -> np.ndarray:
    """
    Recurrently predict concentrations for the forecast horizon.

    Parameters
    ----------
    pipe          : fitted sklearn Pipeline
    df_forecast   : rows with year >= FORECAST_FROM (LK columns already injected)
    seed_storage  : storage_conc from the last training row (seeds the recursion)
    leach_scale   : multiplicative factor on ALL leaching columns (used by MC)
                    This perturbs GM, GK, GL AND LK leaching simultaneously.

    Returns
    -------
    predictions : 1-D array, length = len(df_forecast)
    """
    predictions  = []
    prev_storage = seed_storage

    for _, row in df_forecast.iterrows():
        x = np.array([[
            row["half_year"],
            row["prod_gm"],
            row["leach_gm"] * leach_scale,
            row["prod_gk"],
            row["leach_gk"] * leach_scale,
            row["leach_gl"] * leach_scale,
            row["prod_lk"],
            row["leach_lk"] * leach_scale,   # LK leaching also perturbed in MC
            row["pit_pump"],
            prev_storage,
        ]])
        pred         = pipe.predict(x)[0]
        predictions.append(pred)
        prev_storage = pred     # feed-forward recurrence

    return np.array(predictions)


# ===========================================================================
# 5. MONTE CARLO
# ===========================================================================

def monte_carlo(pipe, df_forecast: pd.DataFrame,
                seed_storage: float,
                rng: np.random.Generator) -> dict:
    """
    N_MC_RUNS forecasts with independent +/- MC_PERTURB uniform leach scaling.
    Returns dict: all_runs, p10, p50, p90
    """
    all_runs = np.zeros((N_MC_RUNS, len(df_forecast)))
    for i in range(N_MC_RUNS):
        scale        = rng.uniform(1.0 - MC_PERTURB, 1.0 + MC_PERTURB)
        all_runs[i]  = run_forecast(pipe, df_forecast, seed_storage, leach_scale=scale)

    return dict(
        all_runs = all_runs,
        p10      = np.percentile(all_runs, 10, axis=0),
        p50      = np.percentile(all_runs, 50, axis=0),
        p90      = np.percentile(all_runs, 90, axis=0),
    )


# ===========================================================================
# 6. PLOT RESULTS
# ===========================================================================

def plot_results(results: dict, output_path: str):
    """
    Thesis-quality 4-panel dark figure:
      - Solid coloured line  : consultant's historical modelled values (2014-2025)
      - Dashed grey line     : consultant's own 2026-2030 extrapolation
      - Solid bright line    : RF P50 forecast
      - Shaded band          : P10-P90 from Monte Carlo
      - Dotted vertical line : train / forecast boundary
    """
    plt.rcParams.update({
        "font.family":       "DejaVu Sans",
        "font.size":         11,
        "axes.titlesize":    13,
        "axes.labelsize":    11,
        "axes.spines.top":   False,
        "axes.spines.right": False,
        "figure.facecolor":  "#0f1117",
        "axes.facecolor":    "#1a1d27",
        "axes.edgecolor":    "#444",
        "axes.labelcolor":   "#e0e0e0",
        "xtick.color":       "#aaa",
        "ytick.color":       "#aaa",
        "text.color":        "#e0e0e0",
        "grid.color":        "#2a2d3a",
        "grid.linewidth":    0.6,
        "axes.grid":         True,
        "legend.framealpha": 0.2,
        "legend.facecolor":  "#1a1d27",
        "legend.edgecolor":  "#444",
    })

    params = list(results.keys())
    fig    = plt.figure(figsize=(16, 11))
    fig.suptitle(
        "Leveaniemi Mine - Process Water Quality Forecast 2026-2030\n"
        "Random Forest with Monte Carlo Uncertainty (P10 / P90, n=300, +/-15%)",
        fontsize=15, fontweight="bold", y=0.98, color="#f0f0f0",
    )

    gs = GridSpec(2, 2, figure=fig,
                  hspace=0.42, wspace=0.32,
                  left=0.07, right=0.97, top=0.91, bottom=0.09)
    pos_map = {
        "Cu":  gs[0, 0],
        "NH4": gs[0, 1],
        "Cl":  gs[1, 0],
        "Ni":  gs[1, 1],
    }

    for param in params:
        ax      = fig.add_subplot(pos_map[param])
        res     = results[param]
        colour  = COLOURS[param]
        cfg     = BLOCKS[param]
        hist    = res["df_train"]
        fcst_df = res["df_forecast"]
        mc      = res["mc"]
        yrs     = fcst_df["year"].values

        # Historical consultant values
        ax.plot(
            hist["year"], hist["modelled_conc"],
            color=colour, lw=2.0, alpha=0.9,
            label="Consultant model (hist.)",
        )

        # Consultant extrapolation 2026-2030
        ax.plot(
            fcst_df["year"], fcst_df["modelled_conc"],
            color="#bbbbbb", lw=1.4, ls="--", alpha=0.7,
            label="Consultant extrap.",
        )

        # P10-P90 band
        ax.fill_between(
            yrs, mc["p10"], mc["p90"],
            color=colour, alpha=0.22,
            label="RF P10-P90",
        )

        # RF P50 forecast
        ax.plot(
            yrs, mc["p50"],
            color=colour, lw=2.5, ls="-", alpha=1.0,
            label="RF P50 forecast",
        )

        # Boundary line
        ax.axvline(x=FORECAST_FROM - 0.5, color="#666", lw=0.9, ls=":")
        ax.text(
            FORECAST_FROM - 0.4, ax.get_ylim()[1],
            "Forecast >>", fontsize=8, color="#888",
            ha="left", va="top",
        )

        # Labels & cosmetics
        unit = cfg["unit"]
        ax.set_title(f"{cfg['label']}  [{unit}]", color="#f0f0f0", pad=6)
        ax.set_xlabel("Year", labelpad=4)
        ax.set_ylabel(f"Concentration  ({unit})", labelpad=4)
        ax.set_xlim(2014, 2031)

        # CV score badge
        cv_r2 = float(np.mean(res["cv_scores"]))
        badge_col = "#44aa66" if cv_r2 >= 0.7 else ("#f0a030" if cv_r2 >= 0.5 else "#ee4444")
        ax.text(
            0.02, 0.97,
            f"CV R2 = {cv_r2:.3f}",
            transform=ax.transAxes, fontsize=9.5, fontweight="bold",
            va="top", ha="left", color=badge_col,
            bbox=dict(boxstyle="round,pad=0.35", fc="#22253a", ec="none", alpha=0.85),
        )

        if param == params[0]:
            ax.legend(loc="upper right", fontsize=8.5, ncol=1)

    # Shared legend
    handles = [
        plt.Line2D([0], [0], color=list(COLOURS.values())[0], lw=2.0,
                   label="Consultant model (hist.)"),
        plt.Line2D([0], [0], color="#bbbbbb", lw=1.4, ls="--",
                   label="Consultant extrapolation (2026-2030)"),
        mpatches.Patch(alpha=0.35, color="#888888",
                       label="RF P10-P90  (MC +/-15%, n=300)"),
        plt.Line2D([0], [0], color="#ffffff", lw=2.5,
                   label="RF P50 forecast"),
    ]
    fig.legend(
        handles=handles, loc="lower center", ncol=4,
        fontsize=9.5, framealpha=0.2,
        facecolor="#1a1d27", edgecolor="#444",
        bbox_to_anchor=(0.5, 0.01),
    )

    plt.savefig(output_path, dpi=180, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f"\n  Figure saved: {output_path}")


# ===========================================================================
# 7. EXPORT EXCEL
# ===========================================================================

def export_excel(results: dict, cv_summary: list, output_path: str):
    """Write forecast tables and CV diagnostics to a formatted workbook."""
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        wb  = writer.book
        hdr = wb.add_format({
            "bold": True, "bg_color": "#1f4e79",
            "font_color": "white", "border": 1, "align": "center",
        })
        num = wb.add_format({"num_format": "0.0000", "border": 1})
        txt = wb.add_format({"border": 1})
        ttl = wb.add_format({
            "bold": True, "font_size": 12, "font_color": "#1f4e79",
        })

        # Per-parameter forecast sheets
        for param, res in results.items():
            mc      = res["mc"]
            fcst_df = res["df_forecast"]
            unit    = BLOCKS[param]["unit"]
            sheet   = param   # all names are safe ASCII

            rows_out = []
            for i, (_, row) in enumerate(fcst_df.iterrows()):
                rows_out.append({
                    "Year":               row["year"],
                    "Season":             "Summer" if int(row["half_year"]) == 1 else "Winter",
                    f"P10 ({unit})":      mc["p10"][i],
                    f"P50 ({unit})":      mc["p50"][i],
                    f"P90 ({unit})":      mc["p90"][i],
                    f"Consultant ({unit})": row["modelled_conc"],
                })

            df_out = pd.DataFrame(rows_out)
            df_out.to_excel(writer, sheet_name=sheet, index=False, startrow=2)

            ws = writer.sheets[sheet]
            ws.write(
                0, 0,
                f"{param} -- ML Forecast 2026-2030 | "
                f"RF + Monte Carlo (n={N_MC_RUNS}, +/-{int(MC_PERTURB*100)}%)",
                ttl,
            )
            for ci, col_name in enumerate(df_out.columns):
                ws.write(2, ci, col_name, hdr)
                ws.set_column(ci, ci, max(16, len(col_name) + 3))
            for ri in range(len(df_out)):
                for ci in range(len(df_out.columns)):
                    v = df_out.iloc[ri, ci]
                    ws.write(ri + 3, ci, v, num if isinstance(v, float) else txt)

        # CV diagnostics sheet
        cv_df = pd.DataFrame(cv_summary)
        cv_df.to_excel(writer, sheet_name="CV_Diagnostics", index=False, startrow=2)
        ws = writer.sheets["CV_Diagnostics"]
        ws.write(0, 0, "Cross-Validation Diagnostics -- Random Forest Models", ttl)
        for ci, col_name in enumerate(cv_df.columns):
            ws.write(2, ci, col_name, hdr)
            ws.set_column(ci, ci, max(18, len(col_name) + 3))

    print(f"  Excel saved:  {output_path}")


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    rng = np.random.default_rng(seed=42)

    print("Loading Excel workbook ...")
    df_raw = pd.read_excel(
        EXCEL_FILE, sheet_name=SHEET_NAME,
        header=None, engine="openpyxl",
    )
    print(f"  Sheet dimensions: {df_raw.shape[0]} rows x {df_raw.shape[1]} cols")

    results    = {}
    cv_summary = []

    for param in BLOCKS:
        print("\n" + "-" * 62)
        print(f"  Parameter: {param}")
        print("-" * 62)

        # Parse
        df = parse_block(df_raw, param)
        print(f"  Parsed {len(df)} rows  ({df['year'].min()} to {df['year'].max()})")

        # Features (adds lagged storage; drops first row)
        df = build_features(df)

        # Split
        df_train    = df[df["year"] <= TRAIN_UNTIL].copy()
        df_forecast = df[df["year"] >= FORECAST_FROM].copy()

        # Inject LK leaching into forecast rows (LK = 0 in training data)
        # 50/50 Leveaniemi+Kiruna rates; production = GK placeholder volumes
        lk_rate = LK_LEACH_RATES[param]
        df_forecast = inject_lk_forecast(df_forecast, param)
        print(f"  Training rows : {len(df_train)}"
              f"  ({df_train['year'].min()} to {df_train['year'].max()})")
        print(f"  Forecast rows : {len(df_forecast)}"
              f"  ({df_forecast['year'].min()} to {df_forecast['year'].max()})")
        print(f"  LK leach rate : {lk_rate:.4f} kg/Mton  "
              f"(prod winter={LK_PROD_WINTER} / summer={LK_PROD_SUMMER} Mton -- PLACEHOLDER)")

        # Train (training data still has prod_lk=leach_lk=0)
        pipe, cv_scores, in_sample_r2 = train_model(df_train, param)

        # Seed the recurrence from last training row's storage
        seed_storage = df_train["storage_conc"].iloc[-1]

        # Monte Carlo -- LK leaching is also perturbed +-15%
        print(f"  Running {N_MC_RUNS} MC runs (incl. LK) ...")
        mc = monte_carlo(pipe, df_forecast, seed_storage, rng)
        print(f"  P50 range: {mc['p50'].min():.4f} -- {mc['p50'].max():.4f}"
              f"  {BLOCKS[param]['unit']}")

        results[param] = dict(
            df_train     = df_train,
            df_forecast  = df_forecast,
            pipe         = pipe,
            cv_scores    = cv_scores,
            in_sample_r2 = in_sample_r2,
            mc           = mc,
        )
        cv_summary.append({
            "Parameter":           param,
            "Unit":                BLOCKS[param]["unit"],
            "Training rows":       len(df_train),
            "Forecast rows":       len(df_forecast),
            "CV R2 mean":          round(float(np.mean(cv_scores)), 4),
            "CV R2 std":           round(float(np.std(cv_scores)), 4),
            "In-sample R2":        round(in_sample_r2, 4),
            "CV folds":            N_CV_SPLITS,
            "MC runs":             N_MC_RUNS,
            "MC perturbation %":   int(MC_PERTURB * 100),
            "LK leach rate (kg/Mton)": round(LK_LEACH_RATES[param], 4),
            "LK prod winter (Mton)":   LK_PROD_WINTER,
            "LK prod summer (Mton)":   LK_PROD_SUMMER,
            "LK assumption":           "50/50 Leveaniemi+Kiruna (PLACEHOLDER)",
        })

    print("\n" + "=" * 62)
    print("  Generating figure ...")
    plot_results(results, FIGURE_OUT)

    print("  Exporting Excel ...")
    export_excel(results, cv_summary, EXCEL_OUT)

    print("\n" + "=" * 62)
    print("  DONE")
    print("=" * 62)
    print(f"  >> {FIGURE_OUT}")
    print(f"  >> {EXCEL_OUT}\n")


if __name__ == "__main__":
    main()
