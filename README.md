# Water Quality Forecasting — Leveäniemi Mine

Machine learning model for predicting process water quality at the Leveäniemi iron ore mine (LKAB, northern Sweden).

## Project Overview

Uses a Random Forest model trained on the consultant's 2013 mass-balance model data to forecast concentrations of Cu, NH₄, Cl, and Ni in process water from 2026–2030, with Monte Carlo uncertainty quantification (P10/P50/P90 bands).

## Files

| File | Description |
|------|-------------|
| `water_quality_model.py` | Main ML script — run this to generate all outputs |
| `PROJECT_DOCUMENTATION.md` | Full documentation (plain language, no ML background needed) |
| `forecast_figure.png` | 4-panel thesis figure (generated output) |
| `forecast_results.xlsx` | Forecast table + CV diagnostics (generated output) |

## Requirements

```bash
pip install pandas numpy scikit-learn matplotlib openpyxl xlsxwriter
```

## Usage

Place the Excel data file in the same directory, then:

```bash
python water_quality_model.py
```

Runtime: ~10–15 minutes. Outputs: `forecast_figure.png` and `forecast_results.xlsx`.

## Results

| Parameter | CV R² | Unit |
|-----------|-------|------|
| Cu | 0.765 | µg/L |
| NH₄ | 0.780 | mg/L |
| Cl | 0.856 | mg/L |
| Ni | 0.857 | µg/L |

## Method

- **Model**: Random Forest (500 trees, StandardScaler pipeline)
- **Validation**: Temporal 4-fold cross-validation
- **Uncertainty**: Monte Carlo simulation, 300 runs, ±15% leaching rate perturbation
- **Forecast horizon**: 2026.0 – 2030.0 (9 half-year periods)

## Master's Thesis

This project is part of a Master's thesis at a Swedish university, supervised analysis of LKAB's Leveäniemi mine water balance model originally developed by an environmental consultant in 2013.
