"""Script to generate the Colab .ipynb notebook from mass_balance_colab.py"""
import json, re

# Read the colab py file
with open("mass_balance_colab.py", "r", encoding="utf-8") as f:
    full_code = f.read()

# ── Cell definitions ─────────────────────────────────────────

cells = []

def md_cell(src):
    return {"cell_type": "markdown", "metadata": {}, "source": src}

def code_cell(src):
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": src,
    }

# Cell 1 – title
cells.append(md_cell([
    "# Leveäniemi Water Quality — Mass-Balance Model\n",
    "**Method 1: Physics-based mass-balance model** for LK discharge water quality.\n\n",
    "- Validation period: 2016–2025 vs SVA79 monitoring data\n",
    "- Forecast period: 2026–2030\n",
    "- Parameters: SO₄, Ca, Cl, NO₃, NH₄, Cu, Ni, Zn, Co, Mo, As, Cr\n\n",
    "> Run all cells top to bottom. Upload `parameters_used2.xlsx` when prompted in **Cell 2**.",
]))

# Cell 2 – install
cells.append(md_cell(["## Cell 1 — Install required packages"]))
cells.append(code_cell(["!pip install openpyxl xlsxwriter --quiet"]))

# Cell 3 – upload
cells.append(md_cell(["## Cell 2 — Upload data file\nA file picker will appear. Upload **parameters_used2.xlsx**."]))
cells.append(code_cell([
    "from google.colab import files\n",
    "uploaded = files.upload()   # upload parameters_used2.xlsx\n",
    "PARAMS_FILE = list(uploaded.keys())[0]\n",
    "print('Using file:', PARAMS_FILE)",
]))

# Cell 4 – imports & constants (everything up to load_data)
cells.append(md_cell(["## Cell 3 — Imports, constants & leaching rates"]))
# Extract just the imports + constants (skip the header comments about how to use)
imports_end = full_code.find("def load_data(")
imports_block = full_code[full_code.find("import warnings"):imports_end].strip()
cells.append(code_cell([imports_block]))

# Cell 5 – load_data function
cells.append(md_cell(["## Cell 4 — Data loading"]))
ld_start = full_code.find("def load_data(")
ld_end   = full_code.find("\ndef _get_vol(")
cells.append(code_cell([full_code[ld_start:ld_end].strip()]))

# Cell 6 – build_inputs + helpers
cells.append(md_cell(["## Cell 5 — Input builder"]))
bi_start = full_code.find("\ndef _get_vol(")
bi_end   = full_code.find("\ndef run_mass_balance(")
cells.append(code_cell([full_code[bi_start:bi_end].strip()]))

# Cell 7 – run_mass_balance
cells.append(md_cell(["## Cell 6 — Mass-balance recurrence"]))
mb_start = full_code.find("\ndef run_mass_balance(")
mb_end   = full_code.find("\ndef plot_results(")
cells.append(code_cell([full_code[mb_start:mb_end].strip()]))

# Cell 8 – plot_results
cells.append(md_cell(["## Cell 7 — Plotting (inline in Colab)"]))
pl_start = full_code.find("\ndef plot_results(")
pl_end   = full_code.find("\ndef export_excel(")
cells.append(code_cell([full_code[pl_start:pl_end].strip()]))

# Cell 9 – export_excel
cells.append(md_cell(["## Cell 8 — Excel export"]))
ex_start = full_code.find("\ndef export_excel(")
ex_end   = full_code.find("\ndef main(")
cells.append(code_cell([full_code[ex_start:ex_end].strip()]))

# Cell 10 – main
cells.append(md_cell(["## Cell 9 — Main function"]))
mn_start = full_code.find("\ndef main(")
mn_end   = full_code.find("\n# Run!")
cells.append(code_cell([full_code[mn_start:mn_end].strip()]))

# Cell 11 – run
cells.append(md_cell(["## Cell 10 — Run the model\nThis will display the 12-parameter chart inline and offer downloads."]))
cells.append(code_cell(["results = main(PARAMS_FILE)"]))

# ── Assemble notebook ─────────────────────────────────────────
nb = {
    "nbformat": 4,
    "nbformat_minor": 4,
    "metadata": {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3",
        },
        "language_info": {
            "name": "python",
            "version": "3.10.0",
        },
        "colab": {
            "provenance": [],
            "name": "Leveaniemi_Water_Quality_Model.ipynb",
        },
    },
    "cells": cells,
}

out_path = "Leveaniemi_Water_Quality_Model.ipynb"
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print(f"Notebook written -> {out_path}")
print(f"Cells: {len(cells)}")
