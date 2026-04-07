# Hart–CADMID–R10 + Fuzzy TOPSIS (3 distances) — reproducible repository

This repository starts **only** from the **linguistic inputs per scenario (A–D)** and reproduces:

1) Fuzzy TOPSIS closeness coefficients and ranks using **3 distance metrics**:
   - Euclidean (L2)
   - Manhattan (L1)
   - Chebyshev (L∞)

2) Hart–CADMID temporal weighting with **3 transition profiles**:
   - Fast (γ = 0.5)
   - Moderate (γ = 1.0)
   - Slow (γ = 2.0)

3) Phase-dependent priorities and **R10 strategy mapping** by CADMID phase.

## Folder structure

- `data/linguistic_inputs.xlsx`  
  The **only input** file: linguistic terms for each indicator under scenarios A–D.
- `scripts/01_generate_topsis_3dist.py`  
  Reads `data/linguistic_inputs.xlsx` and writes `output/topsis_3dist_results.xlsx`.
- `scripts/02_generate_hart_cadmid_r10.py`  
  Reads `output/topsis_3dist_results.xlsx` and writes `output/hart_cadmid_r10_transition.xlsx`.
- `output/`  
  Generated XLSX files.

## How to run

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

python scripts/01_generate_topsis_3dist.py
python scripts/02_generate_hart_cadmid_r10.py
```

## Outputs

- `output/topsis_3dist_results.xlsx`
  - per-dimension CC and ranks for L2/L1/L∞
- `output/hart_cadmid_r10_transition.xlsx`
  - separate sheets for Fast / Moderate / Slow transitions
  - per phase: Hart–CADMID weights and final phase priorities
  - includes assigned R10 codes per phase (compact notation R0..R9)
