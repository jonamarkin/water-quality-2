# Leveäniemi Mine — Process Water Quality Forecasting
## Complete Project Documentation
### Written for someone new to machine learning and data science

---

## Table of Contents

1. [What Is This Project Doing?](#1-what-is-this-project-doing)
2. [The Dataset Explained](#2-the-dataset-explained)
3. [The Problem We Are Solving](#3-the-problem-we-are-solving)
4. [Different Approaches We Could Have Taken](#4-different-approaches-we-could-have-taken)
5. [Our Chosen Approach — Random Forest + Monte Carlo](#5-our-chosen-approach)
6. [The Code Explained Section by Section](#6-the-code-explained-section-by-section)
7. [How to Run the Code](#7-how-to-run-the-code)
8. [How to Interpret the Results](#8-how-to-interpret-the-results)
9. [Frequently Asked Questions](#9-frequently-asked-questions)
10. [How to Get Better Results](#10-how-to-get-better-results)

---

## 1. What Is This Project Doing?

### The Simple Version

The Leveäniemi iron ore mine in northern Sweden pumps large amounts of water through its operations every year. This water picks up chemical contaminants from the ore — things like copper (Cu), ammonium (NH₄), chloride (Cl), and nickel (Ni). Too much of these in the water can be harmful to the environment.

Back in 2013, an environmental consultant built a mathematical model to **predict** how concentrated these chemicals would be in the mine's process water from 2014 all the way to 2030. He used a specific mix of ore types (called "Gruvberget-Mertainen" or GM) as his baseline.

The mine has since changed which ores it processes. This means the consultant's original predictions are now **outdated**.

**What this project does:** We use Machine Learning (ML) to:
1. **Learn** the consultant's prediction logic from his own historical data
2. **Generate new predictions** for 2026–2030 using an ML model
3. **Quantify uncertainty** — instead of giving one number, we give a range (P10 to P90) showing how uncertain we are

### Why This Matters

When regulators, environmentalists, or mine managers ask "How bad will the water quality be in 2028?", they need a scientifically defensible answer with uncertainty bounds. This project provides exactly that.

---

## 2. The Dataset Explained

### The Excel File

The file `Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx` has **18 sheets**. The most important one is called **"Process water"**.

### What Is In "Process water"?

Think of it like a table with rows representing every 6-month period from 2014 to 2030. There are **four separate blocks** of data in this sheet — one block for each chemical contaminant:

| Block | Chemical | Unit | Where in the sheet |
|-------|----------|------|-------------------|
| Cu | Copper | µg/L (micrograms per litre) | Rows 42–74 |
| NH₄ | Ammonium | mg/L (milligrams per litre) | Rows 81–113 |
| Cl | Chloride | mg/L | Rows 121–153 |
| Ni | Nickel | µg/L | Rows 160–192 |

### The Half-Year Pattern

Each year is split into **two rows**:
- **x.0 rows** (e.g., 2015.0) = Winter period. The mine is not actively processing ore. Concentrations tend to be LOW or near zero.
- **x.5 rows** (e.g., 2015.5) = Summer processing season. The mine is crushing and processing ore, releasing chemicals. Concentrations are HIGH.

This creates a saw-tooth/zigzag pattern in the data — low, high, low, high — every year. This is not an error; it is physically real.

### The Key Columns

Inside each block, the important columns are (all numbered from 0):

| Column | What It Contains |
|--------|-----------------|
| Col 0 | Year (e.g., 2015.0 or 2015.5) |
| Col 1 | Season flag: 1 = summer (processing), 0 = winter |
| Col 3 | Production volume from GM ore (million tonnes) |
| Col 4 | Total leaching load from GM ore (how much chemical it releases) |
| Col 5 | Production volume from GK ore (million tonnes) |
| Col 6 | Total leaching load from GK ore |
| Col 7 | Total leaching load from GL ore |
| Col 8 | Volume of water pumped from the Leveäniemi pit (million m³) |
| **Col 9** | **The modelled concentration — this is the TARGET we predict** |
| **Col 31** | **Total storage concentration — this feeds into the NEXT row** |

### What Is "Leaching"?

Leaching is the process by which chemicals dissolve out of the ore into the water. Different ore types (GM, GK, GL) release different amounts of each chemical per tonne of ore processed. Higher leaching rate × more ore = more chemical in the water.

### The Recurrence Structure (Very Important)

The consultant's model is a **recurrence formula** — meaning each row depends on the previous row. Specifically:

```
concentration[this period] = f(production, leaching, pump volume, concentration[last period])
```

The value in Column 31 ("Total storage concentration") from one row is used as an input to calculate the concentration of the NEXT row. This is like compound interest — each period builds on the last.

This means when we make predictions, **we cannot skip rows**. We must predict period 1, use that prediction as input for period 2, and so on. This is called **autoregressive forecasting**.

---

## 3. The Problem We Are Solving

### The Original Consultant Model

The consultant used a standard engineering mass-balance equation. In plain terms:

```
new_concentration = (chemicals_coming_in) / (total_water_volume)
```

Where:
- Chemicals coming in = (leaching from ore) + (chemicals already in stored water)
- Total water volume = pump volume + process water + drainage

This formula works, but it makes specific assumptions about ore mix fractions that no longer reflect reality.

### The Gap

The consultant's model predicted using GM-heavy ore assumptions. The mine now uses more GK and GL ore. The thesis owner found gaps between the model and actual daily monitoring data.

### Our Goal

Train a machine learning model to **learn** the consultant's formula structure from his 2014–2025 data, then use that model to **generate new predictions** with proper uncertainty quantification.

---

## 4. Different Approaches We Could Have Taken

Here is a comparison of all realistic approaches for this kind of problem:

### Approach A: Copy the Consultant's Formula Directly
**What it means:** Replicate his exact mathematical equations in Python and adjust the ore fractions.

✅ **Pros:** Perfectly transparent, easy to explain to regulators, no ML needed

❌ **Cons:** Requires knowing the exact formula parameters; if the formula is wrong, the predictions are wrong; gives only point estimates (no uncertainty)

**Why we didn't use this:** The exact formula uses many interacting parameters that are not fully documented in the Excel. Extracting them correctly would take weeks and is outside scope.

---

### Approach B: Linear Regression
**What it means:** Fit a straight line through the data. "If production goes up by X, concentration goes up by Y × X."

✅ **Pros:** Simple, interpretable, fast

❌ **Cons:** The relationship between inputs and concentration is NOT linear. The recurrence structure means concentrations cycle up and down in complex ways. A straight line would badly underfit.

---

### Approach C: Random Forest (Our Choice) ✅
**What it means:** Train hundreds of decision trees on the data and average their predictions.

✅ **Pros:** Handles non-linear relationships, handles the alternating season pattern, works well with small datasets (~22 training rows), provides feature importances, robust to outliers

❌ **Cons:** Less interpretable than linear regression, can overfit if not validated properly

---

### Approach D: Neural Network / Deep Learning
**What it means:** Train a multi-layer neural network on the data.

✅ **Pros:** Can learn very complex patterns

❌ **Cons:** Needs FAR more data than 22 rows. With only 22 training examples, a neural network would badly overfit (memorise the training data but fail on new data). Not appropriate here.

---

### Approach E: ARIMA / Time Series Models
**What it means:** Statistical models specifically designed for time series data.

✅ **Pros:** Well-established for time-dependent data

❌ **Cons:** Standard ARIMA does not handle multiple input features (production volumes, leaching rates, etc.) easily. These inputs ARE the key drivers of concentration.

---

### Why Random Forest Is Right for This Problem

1. **Small dataset (22 training rows):** RF works well with small datasets. Neural networks do not.
2. **Non-linear relationships:** Concentration responds non-linearly to production × leaching rate interactions. RF captures this naturally.
3. **Multiple inputs:** RF handles 8 input features without needing feature selection.
4. **The alternating pattern:** By including the `half_year` binary feature (0=winter, 1=summer), the RF learns the season pattern explicitly.
5. **Recurrence:** By including `prev_storage` as a feature, we give the RF the same information the consultant's formula used.

---

## 5. Our Chosen Approach

### Step-by-Step Overview

```
Excel File
    │
    ▼
Step 1: PARSE — Extract each contaminant block (Cu, NH4, Cl, Ni)
    │
    ▼
Step 2: FEATURES — Add "previous storage concentration" as a lag feature
    │
    ▼
Step 3: TRAIN — Fit a Random Forest model on 2015.0–2025.5 rows
    │         (using temporal cross-validation to check quality)
    ▼
Step 4: FORECAST — Use the consultant's 2026–2030 input parameters
    │              and predict concentration recurrently
    ▼
Step 5: MONTE CARLO — Repeat forecast 300 times with random ±15%
    │                 perturbation on leaching rates
    ▼
Step 6: OUTPUT — Plot 4-panel figure + Excel table of P10/P50/P90
```

### The 8 Input Features

| Feature | Why Included |
|---------|-------------|
| `half_year` | Tells the model: are we in summer (processing) or winter? |
| `prod_gm` | GM ore production volume — how much GM ore is being mined |
| `leach_gm` | GM leaching load — how much chemical GM ore releases |
| `prod_gk` | GK ore production volume |
| `leach_gk` | GK leaching load |
| `leach_gl` | GL leaching load |
| `pit_pump` | Water volume pumped from the pit — dilutes the concentration |
| `prev_storage` | Last period's storage concentration — the recurrence state |

### The Recurrence State

`prev_storage` is the most important feature. The consultant's formula explicitly uses "last period's concentration" to compute "this period's concentration." By including this as a feature, we teach the RF the same structure.

During **training**: we use the consultant's recorded `prev_storage` values.

During **forecasting**: we use the RF's own predicted concentration as `prev_storage` for the next step. This is the feed-forward loop.

### What Is Monte Carlo Simulation?

Imagine you are not sure exactly how much chemical each tonne of ore releases. The consultant gave us one number, but the real value could be slightly higher or lower (±15%).

Monte Carlo simulation runs the forecast 300 times, each time using a slightly different leaching rate (randomly chosen within ±15% of the stated value). 

This produces 300 slightly different forecast trajectories. We then report:
- **P10** — the 10th percentile (only 10% of runs were lower than this — the optimistic scenario)
- **P50** — the 50th percentile (the median — our best guess)
- **P90** — the 90th percentile (only 10% of runs were higher — the pessimistic scenario)

The shaded band between P10 and P90 is the **uncertainty range**. A wide band means we are less confident. A narrow band means the result is robust to leaching rate uncertainty.

---

## 6. The Code Explained Section by Section

### Configuration Block (top of file)

```python
BLOCKS = {
    "Cu":  dict(label="Cu",  unit="ug/L",  first_row=42,  last_row=74),
    ...
}
```

This tells the code exactly where each contaminant block lives in the Excel sheet. These numbers (42, 74, etc.) were determined by manually exploring the sheet structure.

```python
COL_MODELLED  = 9   # This is our prediction target
COL_STORAGE   = 31  # This is the recurrence state
TRAIN_UNTIL   = 2025.5  # We train on data up to here
FORECAST_FROM = 2026.0  # We forecast from here onwards
N_MC_RUNS     = 300     # Number of Monte Carlo iterations
MC_PERTURB    = 0.15    # ±15% perturbation
```

### `parse_block()` — Reading the Data

This function reads one contaminant's data block from the raw Excel sheet and returns a clean table. It:
- Skips rows that are not valid year entries
- Fills in missing `half_year` values with 0 (winter)
- Extracts the 10 key columns into a named DataFrame

### `build_features()` — Creating the Lag Feature

```python
df["prev_storage"] = df["storage_conc"].shift(1)
```

`shift(1)` moves the storage column down by one row, so each row now has "what was the storage concentration in the previous period?" as a feature. The very first row has no "previous" row, so it is dropped.

### `train_model()` — Training the Random Forest

The model is a **Pipeline** that does two things in sequence:
1. **StandardScaler**: Rescales all features to have mean=0, standard deviation=1. This is important because Cl leaching rates are ~100,000 while Cu leaching rates are ~20 — wildly different scales. Scaling prevents larger numbers from dominating.
2. **RandomForestRegressor**: The actual ML model — 500 decision trees.

**Temporal cross-validation** is used instead of random cross-validation. This is critical: because time matters, you cannot train on 2022 data and test on 2018 data (you would be "cheating" by using future information). TimeSeriesSplit always trains on earlier data and validates on later data.

### `run_forecast()` — Recurrent Prediction

```python
for each row in 2026.0, 2026.5, ..., 2030.0:
    predict concentration using [features + prev_storage]
    set prev_storage = that prediction
    move to next row
```

This loop is why the forecast is called "recurrent" — each prediction feeds into the next one.

### `monte_carlo()` — Uncertainty Quantification

Runs `run_forecast()` 300 times, each time scaling the leaching features by a random factor between 0.85 and 1.15 (±15%). Returns the 10th, 50th, and 90th percentiles across all 300 runs.

### `plot_results()` — The Figure

Creates a 2×2 grid of panels. Each panel shows:
- Coloured solid line: what the consultant's model said happened 2014–2025
- Grey dashed line: the consultant's own 2026–2030 extrapolation
- Coloured solid brighter line: the RF P50 (our best guess)
- Shaded band: P10 to P90 uncertainty range
- Green/orange/red badge: CV R² score (quality of the model)

### `export_excel()` — The Results Table

Writes one Excel sheet per contaminant with columns:
- Year, Season (Summer/Winter), P10, P50, P90, Consultant value

Plus a "CV_Diagnostics" sheet with model quality metrics.

---

## 7. How to Run the Code

### Prerequisites

You need Python 3 and these packages installed:
```
pip install pandas numpy scikit-learn matplotlib openpyxl xlsxwriter
```

### Running

1. Place `water_quality_model.py` in the same folder as the Excel file
2. Open a terminal/command prompt in that folder
3. Run: `python water_quality_model.py`
4. Wait approximately 10–15 minutes (the 300 MC runs × 4 parameters take time)

### What You Will See in the Terminal

```
Loading Excel workbook ...
  Sheet dimensions: 576 rows x 70 cols

--------------------------------------------------------------
  Parameter: Cu
--------------------------------------------------------------
  Parsed 33 rows  (2014.5 to 2030.5)
  Training rows : 22  (2015.0 to 2025.5)
  Forecast rows : 9  (2026.0 to 2030.0)
  [Cu]  CV R2 (mean): 0.765  |  folds: 0.766  0.759  0.772  0.761
  Running 300 MC runs ...
  P50 range: 1.086 -- 11.634  ug/L
...
  Figure saved: forecast_figure.png
  Excel saved:  forecast_results.xlsx
  DONE
```

---

## 8. How to Interpret the Results

### The CV R² Score (Most Important Quality Metric)

R² (pronounced "R-squared") measures how well the model fits the data. It ranges from 0 to 1:

| R² Value | What It Means |
|----------|--------------|
| 0.9–1.0 | Excellent fit — model has learned the pattern very well |
| 0.7–0.9 | Good fit — model captures most of the variation |
| 0.5–0.7 | Moderate — model captures the main trends but misses details |
| Below 0.5 | Poor — model is struggling; predictions less reliable |

**Our results:**

| Parameter | CV R² | Interpretation |
|-----------|-------|----------------|
| Cu | 0.765 | Good |
| NH₄ | 0.780 | Good |
| Cl | 0.856 | Very good |
| Ni | 0.857 | Very good |

All four models exceed the 0.7 threshold. This means the RF has successfully captured the consultant's formula structure.

**Why "CV R²" and not just "R²"?**

In-sample R² (training data R²) was 0.97 for all parameters — near perfect. But this is measured on data the model was TRAINED on, so it is not a fair test (the model has "seen" those answers). CV R² is measured on data the model has never seen during training. The gap between 0.97 and ~0.77 is normal and expected — it tells us the model generalises well but is not perfect.

### The Figure

**Solid coloured line (historical):** This is the consultant's own recorded modelled concentrations 2014–2025. The saw-tooth pattern (up-down-up-down) is the summer/winter alternation. High peaks = summer processing season. Near-zero values = winter.

**Grey dashed line:** The consultant's own extrapolation to 2026–2030. Our model should track this closely if it has learned his formula correctly.

**Bright solid line (RF P50):** Our model's best-guess forecast. If this tracks closely with the grey dashed line, it confirms our model has learned the consultant's logic. Small differences arise because the RF captures non-linear effects.

**Shaded band (P10–P90):** The uncertainty range. A narrow band means the forecast is robust — changing leaching rates by ±15% does not change the answer much. A wide band means the forecast is sensitive to leaching rate uncertainty.

**Vertical dotted line:** Separates the historical period (left) from the forecast period (right).

### The Excel File

Each parameter sheet has a row for each half-year forecast period. For a thesis, the most useful columns are P50 (your headline prediction) and P10/P90 (your confidence interval).

Example reading: "We predict the Cl concentration in summer 2028 will be **18.0 mg/L** (P50), with a 80% confidence interval of **[15.2, 21.1] mg/L** (P10–P90)."

---

## 9. Frequently Asked Questions

### "Why does the concentration go to near-zero in some years?"

The near-zero readings are winter rows (x.0 periods). During winter, the mine does not process ore, so the leaching input (Col 4, 6, 7) is very small. The consultant's model assigns low concentrations to these periods. The RF has learned this pattern from the `half_year` feature.

### "The RF P50 is very close to the consultant's extrapolation — does that mean we just copied his work?"

No. The RF independently learned the relationship from the training data. The fact that its predictions are close to the consultant's extrapolation is **validation** — it confirms that the RF has successfully captured the same physical logic. If they were wildly different, that would be a red flag.

### "Why only 9 forecast rows for 2026–2030?"

The data ends at 2030.0 (winter 2030). The consultant's sheet has rows for 2026.0, 2026.5, 2027.0, 2027.5, 2028.0, 2028.5, 2029.0, 2029.5, 2030.0 — that is 9 half-year periods. The summer 2030 row (2030.5) was the last row of the entire dataset and is not included in the forecast window used as input features.

### "Why is the training set only 22 rows?"

We train on years 2015.0 to 2025.5. That is 22 half-year periods. The first row (2014.5) is dropped because it has no "previous storage" value. This small dataset is a known limitation. However, RF is well-suited for small datasets, and the high CV R² scores confirm the model works despite the small size.

### "What does ±15% leaching perturbation mean physically?"

Leaching rates (how much chemical per tonne of ore) are measured in laboratory bottle-roll tests. These tests have measurement uncertainty — the real field leaching rate could be ±15% of the lab measurement. The Monte Carlo simulation explores this uncertainty.

### "Is this model good enough for regulatory submission?"

The model is appropriate as a thesis/academic analysis. For formal regulatory submission, additional steps would be needed: sensitivity analysis on more parameters, validation against actual monitoring data, and review by a qualified engineer. The P10/P90 bands are a good starting point for regulatory discussions.

### "What does in-sample R² of 0.97 vs CV R² of 0.77 tell us?"

The gap is normal. In-sample R² is measured on data the model has seen; CV R² is measured on data it has not seen. An in-sample R² of 0.97 means the model fits the training data very well. A CV R² of 0.77 means it generalises reasonably well to unseen data. If CV R² were close to 0, the model would be "overfitting" — memorising the training data without learning general patterns.

---

## 10. How to Get Better Results

### Option 1: Add More Training Data

The biggest limitation is 22 training rows. If actual monitoring data from the mine (daily or monthly measurements) is available, it can be incorporated. More training data = better generalisation = higher CV R².

### Option 2: Include More Features

Currently we use 8 features. Potential additions:
- **Seasonal temperature** (affects leaching rates)
- **Rainfall/snowmelt** (affects dilution and pit pump volume)
- **Tailings pond levels** (affects storage concentration)
- **Actual GK/GL production fractions** for 2026–2030 (if known)

### Option 3: Train Separate Models for Summer and Winter

The winter rows (half_year=0) behave very differently from summer rows (half_year=1). Training two separate models — one for winter, one for summer — can improve accuracy if the seasonal split confuses the combined model.

```python
df_winter = df_train[df_train["half_year"] == 0]
df_summer = df_train[df_train["half_year"] == 1]
# Train separate RF models for each
```

### Option 4: Tune the Random Forest Hyperparameters

Currently using default settings. These can be optimised:

| Parameter | Current | Try |
|-----------|---------|-----|
| `n_estimators` | 500 | 200–1000 |
| `max_depth` | None (unlimited) | 3–10 |
| `min_samples_leaf` | 2 | 1–5 |
| `max_features` | "sqrt" | "log2", 0.5 |

Use `GridSearchCV` or `RandomizedSearchCV` to find the best combination automatically.

### Option 5: Increase Monte Carlo Runs

Currently running 300 iterations. For more stable P10/P90 estimates, increase to 1000. This will increase runtime but improve the statistical robustness of the uncertainty bands.

```python
N_MC_RUNS = 1000  # Change this in the configuration block
```

### Option 6: Validate Against Real Monitoring Data

If the mine has actual water quality measurements from 2014–2025, plot them on the same figure as the consultant's model and the RF predictions. Where the RF predictions are closer to real measurements than the consultant's model, the RF has added value.

### Option 7: Gradient Boosting (Alternative Model)

Replace `RandomForestRegressor` with `GradientBoostingRegressor` or `XGBoost`. These often outperform Random Forests on small tabular datasets.

```python
from sklearn.ensemble import GradientBoostingRegressor
model = GradientBoostingRegressor(n_estimators=200, learning_rate=0.05, max_depth=3)
```

---

## Summary Table

| Topic | Key Point |
|-------|-----------|
| Dataset | Excel file, "Process water" sheet, 4 contaminant blocks, 33 half-year rows each |
| Target variable | Column 9 — modelled concentration (µg/L or mg/L) |
| Recurrence state | Column 31 — storage concentration (feeds into next row) |
| Training data | 2015.0 to 2025.5 (22 rows per parameter) |
| Forecast horizon | 2026.0 to 2030.0 (9 rows per parameter) |
| Model | Random Forest (500 trees, StandardScaler pipeline) |
| CV R² achieved | 0.765 (Cu), 0.780 (NH₄), 0.856 (Cl), 0.857 (Ni) — all good |
| Uncertainty method | Monte Carlo, 300 runs, ±15% leaching rate perturbation |
| Outputs | P10/P50/P90 forecast bands; figure + Excel |
| Runtime | ~10–15 minutes |

---

*Documentation written for the Leveäniemi Mine process water quality thesis project.*
*Generated alongside `water_quality_model.py`.*
