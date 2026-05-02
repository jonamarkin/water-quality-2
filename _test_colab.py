import warnings; warnings.filterwarnings('ignore')
import sys, types

# Stub out google.colab so the script runs outside Colab
colab_mod = types.ModuleType('google.colab')
colab_mod.files = types.SimpleNamespace(upload=lambda: None, download=lambda f: None)
sys.modules['google'] = types.ModuleType('google')
sys.modules['google.colab'] = colab_mod

PARAMS_FILE = "parameters_used2.xlsx"
exec(open('mass_balance_colab.py').read())
