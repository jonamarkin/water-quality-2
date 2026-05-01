# The Consultant's Mathematical Model — Complete Reconstruction
## Leveäniemi Mine Process Water Quality Model (2013)
### Reverse-engineered from Excel formulas in `Leveäniemi vattenbalans 301013-3-Update-2026Feb26.xlsx`

---

## Overview

The consultant built a **recurrent mass-balance model** to predict the concentration of chemical contaminants in the mine's process water from 2014 to 2030. The model is implemented as a chain of Excel formulas where each row (half-year period) depends on the row immediately above it.

The model is **not a single equation** — it is a system of coupled equations, one per contaminant (Cu, NH₄, Cl, Ni, Zn, Co, Mo, SO₄, Ca, NO₃-N, PO₄-P, F, As, Cr). All contaminants follow **exactly the same mathematical structure**, differing only in their leaching rate constants.

---

## Time Structure

The model runs in **half-year time steps**:

| Row type | Year label | Months | Represents |
|----------|-----------|--------|-----------|
| **x.0** (e.g. 2015.0) | Winter | 5 months (Oct–Feb) | Pit dewatering only, no ore processing |
| **x.5** (e.g. 2015.5) | Summer | 7 months (Mar–Sep) | Active ore processing season |

This alternation repeats every year from 2014.5 to 2030.5 (33 rows per parameter block).

---

## Column Definitions (from Excel header rows)

The following column mapping applies to **all contaminant blocks**:

| Excel Col | Label | Symbol | Units | Description |
|-----------|-------|--------|-------|-------------|
| A | Year | t | — | Year (x.0 = winter, x.5 = summer) |
| B | Season flag | s | 0 or 1 | 0 = winter (no processing), 1 = summer |
| C | Leach rate (GM) | λ_GM | kg/Mton | Per-tonne leaching rate for GM ore |
| D | Production GM | P_GM | Mton | Ore production volume — GM ore |
| E | Leaching load GM | L_GM | kg | Total leaching from GM ore = λ_GM × P_GM |
| F | Production GK | P_GK | Mton | Ore production volume — GK ore |
| G | Leaching load GK (blended) | L_GK | kg | Blended GK+GM leach load |
| H | Leaching load alt. | L_H | kg | Alternative blending scenario |
| I | Pit pump volume | Q_pit | Mm³ | Volume pumped from Leveäniemi pit (from `Modell During mining` sheet) |
| J | Pit pump concentration | C_pit | mg/L (or µg/L) | Contaminant conc. in pit water (from `Modell During mining` sheet) |
| K | Surface water volume | Q_surf | Mm³ | Surface water inflow |
| L | Surface water conc. | C_surf | mg/L | Contaminant conc. in surface water |
| M | Tailings drainage vol. | Q_tail | Mm³ | Water from tailings |
| N | Tailings drainage conc. | C_tail | mg/L | Contaminant conc. in tailings drainage |
| O | Gruvberget inflow vol. | Q_GB | Mm³ | Water from Gruvberget pit area |
| P | Gruvberget conc. | C_GB | mg/L | Contaminant conc. from Gruvberget |
| Q | Discharge volume | Q_out | Mm³ | Volume discharged (leaves system) |
| R | (unused in main formula) | — | — | — |
| S | Process plant loss | Q_loss | Mm³ | Water lost at process plant |
| T | Previous storage conc. | C_{t-1} | mg/L | Storage conc. **from previous period** (recurrence) |
| U | Process water vol. | Q_proc | Mm³ | Volume of process water (25.21 Mm³ constant) |
| AF | Storage concentration | C_store | mg/L | **OUTPUT: total storage concentration this period** |
| AG | Total outflow volume | Q_total_out | Mm³ | Sum of all outgoing water streams |
| AH | Total inflow volume | Q_total_in | Mm³ | Sum of all incoming water streams |
| AI | **Modelled concentration** | **C_mod** | mg/L | **PRIMARY OUTPUT: modelled concentration** |

---

## The Core Recurrence Formula

This is the **exact Excel formula** used in every data row (shown here for the Ni block, row 163):

```excel
AI163 = ((B162 + E162 + I162*J162 + K162*L162 + M162*N162 + O162*P162 - S162*T162)
         / (I162 + K162 + M162 + O162 + Q162 + I162 - S162)
         * AH162 + AE162*AF162)
        / (AH162 + AE162)
```

Note: This formula is **in row 163** but references **row 162** (the previous period). This is the recurrence.

### Translating to Mathematical Notation

Let subscript `t` denote the current period and `t-1` the previous period.

Define the **total leaching input** for period `t-1`:

```
Leach_in(t-1) = B(t-1) + E(t-1) + I(t-1)·J(t-1) + K(t-1)·L(t-1) + M(t-1)·N(t-1) + O(t-1)·P(t-1) - S(t-1)·T(t-1)
```

In words:
```
Leach_in = ore_leaching_GM
         + ore_leaching_GK
         + pit_pump_volume × pit_pump_concentration
         + surface_water_volume × surface_water_concentration
         + tailings_drainage_volume × tailings_concentration
         + Gruvberget_inflow_volume × Gruvberget_concentration
         - process_plant_loss_volume × previous_storage_concentration
```

Define the **total dilution volume** for period `t-1`:

```
Vol_denom(t-1) = I(t-1) + K(t-1) + M(t-1) + O(t-1) + Q(t-1) + I(t-1) - S(t-1)
```

Note: `I` (pit pump) appears **twice** in the denominator — this is intentional in the consultant's formula and appears to account for the pit pump volume both as an inflow and as part of the total water balance.

The **average concentration in the active system** during period `t-1`:

```
C_active(t-1) = Leach_in(t-1) / Vol_denom(t-1)
```

The **storage concentration** at the end of period `t-1` is the weighted average of:
- The active system concentration (weighted by total inflow volume AH)  
- The tailings storage concentration (weighted by tailings pond volume AE)

```
C_store(t-1) = (C_active(t-1) × AH(t-1) + AE(t-1) × C_store_pond(t-1))
               / (AH(t-1) + AE(t-1))
```

The **modelled concentration** at time `t` (column AI, the primary output):

```
C_mod(t) = (Leach_in(t-1) / Vol_denom(t-1) × AH(t-1) + AE(t-1) × AF(t-1))
           / (AH(t-1) + AE(t-1))
```

Or equivalently, written as one equation:

```
                         ⎡ B(t-1) + E(t-1) + I(t-1)·J(t-1) + K(t-1)·L(t-1) + M(t-1)·N(t-1) + O(t-1)·P(t-1) - S(t-1)·C_store(t-2) ⎤
                         ⎢ ──────────────────────────────────────────────────────────────────────────────────────────────────────────── ⎥ × AH(t-1)  +  AE(t-1) × C_store(t-2)
                         ⎣           I(t-1) + K(t-1) + M(t-1) + O(t-1) + Q(t-1) + I(t-1) - S(t-1)                                    ⎦
C_mod(t) = ──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
                                                                     AH(t-1) + AE(t-1)
```

---

## The Leaching Load Formulas (Columns E and G)

### Column E — GM Ore Leaching Load
```excel
E(t) = E$row_header × D(t)
     = λ_GM × P_GM(t)
```
Where `λ_GM` is the leaching rate constant from the header row (e.g. 9.8 kg/Mton for Cu, 110,889 kg/Mton for Cl).

### Column G — GK Ore Leaching Load (Blended with GM)
```excel
G(t) = C(t)×0.4×D(t) + G$row_header×0.6×F(t)
     = λ_GM_field × 0.4 × P_GM(t) + λ_GK × 0.6 × P_GK(t)
```
This is a **40/60 weighted blend**: 40% contribution from GM ore, 60% from GK ore.

### Column D — Production Volume (Seasonal)
```excel
D(t) = 5/12 × 5    [winter, x.0 rows]  → 2.0833 Mton
D(t) = 5.5/12 × 7  [summer, x.5 rows]  → 3.2083 Mton
```
This encodes the **seasonal production schedule**: 5 months of winter at 5 Mton/year rate, and 7 months of summer at 5.5 Mton/year rate.

---

## The Pit Pump Inputs (from `Modell During mining` sheet)

Columns I (volume) and J (concentration) in the Process water sheet are **linked from the `Modell During mining` sheet**:

```excel
I(t) = 'Modell During mining'!$G{row} / 1000000    [Mm³]
J(t) = 'Modell During mining'!$D{row}               [mg/L or µg/L]
```

The `Modell During mining` sheet models the **groundwater and pit lake dynamics** separately, tracking how the pit fills with water as mining progresses and the chemical concentration of that water. Key data from this sheet:

| Year | Pit volume pumped (Mm³) | Concentration (normalized) |
|------|------------------------|---------------------------|
| 2015 | 1.017 | 1.444 (relative) |
| 2020 | 1.525 | 0.988 |
| 2025 | 2.034 | 0.760 |
| 2030 | 2.542 | 0.623 |

The **pit pump volume increases every year** as the pit deepens and more groundwater inflows. The **pit water concentration decreases** as the pit fills with cleaner water (dilution effect).

---

## The Recurrence State (Column AF / T)

Column AF stores the **total storage concentration** — the key state variable that makes this a recurrent model.

For each period, `T(t) = AF(t)` which equals the **previous period's storage concentration** `AF(t-1)`. This value is used in the next row's formula as:
- A **negative term** in the numerator (process plant losses carry out mass at the stored concentration)
- A **weight** in the storage pond averaging

In Excel:
```excel
T(current row) = AF(previous row)
```

In the first row (2014.5), the initial storage concentration `T` is set to the **consultant's assumed initial condition** — typically a value close to background water quality.

---

## Leaching Rate Constants

From the `ore-tail leach calc` sheet (lab bottle-roll test results):

| Contaminant | GM (Grvbg+Mertainen) | GL (Leveäniemi 75µm) | GK (Kiruna) | LK* (50/50 avg) |
|-------------|---------------------|---------------------|------------|----------------|
| Cu | 9.58 kg/Mton | 1.36 kg/Mton | 11.67 kg/Mton | **6.51 kg/Mton** |
| NH₄ | 598.5 kg/Mton | n/a | 3047.3 kg/Mton | **1523.6 kg/Mton** |
| Cl | 110,889 kg/Mton | 154,444 kg/Mton | 285,406 kg/Mton | **219,925 kg/Mton** |
| Ni | 0.435 kg/Mton | 0.631 kg/Mton | 1.174 kg/Mton | **0.903 kg/Mton** |
| Zn | 9.43 kg/Mton | 5.31 kg/Mton | 150.0 kg/Mton | — |
| Co | 0.068 kg/Mton | 0.107 kg/Mton | 0.283 kg/Mton | — |
| Mo | 20.2 kg/Mton | 23.0 kg/Mton | 133.6 kg/Mton | — |
| SO₄ | 269,333 kg/Mton | 260,000 kg/Mton | 1,653,906 kg/Mton | — |
| Ca | 83,773 kg/Mton | 141,778 kg/Mton | 581,426 kg/Mton | — |

*LK = Leveäniemi-Kiruna, not in consultant's original model. Added in this thesis work using 50/50 average.

---

## Fixed Volume Parameters

These are constants held fixed throughout the model (from `Flow-Production` sheet):

| Parameter | Symbol | Value | Units |
|-----------|--------|-------|-------|
| Process water volume | Q_proc | 25.21 | Mm³/year |
| Processing plant loss | Q_loss | 0.36 | Mm³/year |
| Gruvberget inflow | Q_GB | 1.26 | Mm³/year |
| Tailings storage | Q_tail | 0.11 | Mm³/year |
| Surface water inflow | Q_surf | various | Mm³/year |

---

## Step-by-Step Walkthrough of One Period

**Given:** We are computing the modelled Cu concentration for **summer 2016 (year 2016.5)**

**Step 1 — Look up the previous period (winter 2016.0):**
- Production GM: `P_GM = 2.0833 Mton`  
- Leaching rate GM: `λ_GM = 9.58 kg/Mton`
- GM leaching load: `E = 9.58 × 2.0833 = 19.96 kg`
- GK production: `P_GK = 1.6667 Mton`
- GK blended load: `G = 0.44×0.4×2.0833 + 11.7×0.6×1.6667 = 0.366 + 11.7 = 12.07 kg`
- Pit pump volume: `I = 0.848 Mm³`
- Pit concentration: `J = 0.000295 µg/L`
- Surface water: `K×L = 0 × 0 = 0`
- Tailings: `M×N = 0 × 0 = 0`
- Gruvberget: `O×P = 1.26 × 0.19 = 0.239`
- Process loss: `S×T = 0.36 × previous_storage`

**Step 2 — Compute total leaching input:**
```
Leach_in = 0 + 19.96 + 0.848×0.000295 + 0 + 0 + 0.239 - 0.36×C_store(prev)
         = 20.199 - 0.36×C_store(prev)   [µg·Mm³/L = kg equivalent]
```

**Step 3 — Compute total dilution volume:**
```
Vol_denom = 0.848 + 0 + 0 + 1.26 + 0 + 0.848 - 0.36 = 2.596 Mm³
```

**Step 4 — Compute active system concentration:**
```
C_active = Leach_in / Vol_denom
```

**Step 5 — Compute storage-weighted average:**
```
C_mod = (C_active × Q_total_in + Q_tailings_pond × C_tailings_pond)
        / (Q_total_in + Q_tailings_pond)
```

**Step 6 — This becomes the modelled Cu concentration for summer 2016, and feeds into winter 2017 as the recurrence state.**

---

## What the Consultant Assumed (Key Assumptions)

1. **Ore mix fixed at GM/GK/GL:** The consultant used only these three ore types. LK (Leveäniemi-Kiruna) was not considered.

2. **Constant production rates:** `D(t)` uses fixed seasonal schedules (5 months winter × 5 Mton/yr rate; 7 months summer × 5.5 Mton/yr rate).

3. **Leaching rates are constants:** The `λ` values from lab bottle-roll tests are applied uniformly throughout 2014–2030. In reality, leaching rates can vary with ore quality and processing conditions.

4. **40/60 ore blend (col G):** The GK leaching load formula `0.4×GM + 0.6×GK` assumes a specific ore mix ratio. This was the consultant's 2013 assumption and may not reflect actual mine operations.

5. **Pit pump volume is externally modelled:** The pit pump volume and concentration come from the separate `Modell During mining` sheet which models the groundwater system independently.

6. **Volume parameters are constant:** Q_proc (25.21 Mm³), Q_GB (1.26 Mm³), Q_loss (0.36 Mm³) are held fixed for the entire 2014–2030 period.

7. **No seasonal variation in background concentrations:** L, N, P (surface water, tailings, Gruvberget concentrations) are either fixed constants or zero.

---

## Why This Thesis Work Improves on the Consultant's Model

| Limitation | Consultant's approach | This thesis |
|-----------|----------------------|-------------|
| Ore mix | Fixed GM/GK/GL with 40/60 blend assumption | RF learns actual mix effects from data; LK added |
| Uncertainty | Single point estimate | Monte Carlo P10/P50/P90 bands |
| Leaching variability | Fixed constants | ±15% MC perturbation on all rates |
| Validation | No cross-validation | 4-fold temporal CV, CV R² reported |
| New ore types | Not considered | LK (Leveäniemi-Kiruna) incorporated |
| Method transparency | Black-box Excel formulas | Fully documented, reproducible Python code |

---

## Summary Formula (Plain Language)

> **The concentration of a contaminant in the mine's process water during any given half-year period equals: the total mass of that contaminant entering the water system (from ore leaching + groundwater inflow + surface water + tailings drainage, minus what leaves with process plant losses), divided by the total volume of water in the system, averaged with the concentration already stored in the tailings pond from the previous period.**

This is fundamentally a **dilution + accumulation** model: chemicals accumulate in the tailings pond over time, and each period's concentration depends on both fresh inputs and the legacy of previous periods.

---

*Formulas extracted directly from Excel cells using `openpyxl` (data_only=False).*  
*Cross-referenced with the `ore-tail leach calc` sheet for leaching rate constants.*  
*Generated as part of the Leveäniemi Mine water quality forecasting thesis project.*
