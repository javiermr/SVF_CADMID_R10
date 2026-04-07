#!/usr/bin/env python3
"""
Generate fuzzy TOPSIS rankings using 3 distance metrics (L2, L1, Linf)
from the INPUT linguistic evaluations stored in:
  data/linguistic_inputs.xlsx

Only the input columns are used:
  Dimension, Indicator, Scenario A, Scenario B, Scenario C, Scenario D

Outputs:
  output/topsis_3dist_results.xlsx
"""
from __future__ import annotations

import os
import math
import pandas as pd
import numpy as np

TFN_MAP = {
    "L":  (0.0, 0.0, 3.0),
    "LM": (0.0, 3.0, 5.0),
    "M":  (3.0, 5.0, 7.0),
    "HM": (5.0, 7.0, 10.0),
    "H":  (7.0, 10.0, 10.0),
}

SCENARIO_WEIGHTS = {"A": 0.20, "B": 0.40, "C": 0.10, "D": 0.30}

FPIS = (10.0, 10.0, 10.0)
FNIS = (0.0, 0.0, 0.0)

def _norm_term(x: str) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip().upper()

def tfn(term: str) -> tuple[float, float, float]:
    term = _norm_term(term)
    if term not in TFN_MAP:
        raise ValueError(f"Unknown linguistic term: '{term}' (expected one of {list(TFN_MAP)})")
    return TFN_MAP[term]

def agg_tfn(a: str, b: str, c: str, d: str) -> tuple[float, float, float]:
    ta, tb, tc, td = tfn(a), tfn(b), tfn(c), tfn(d)
    wA, wB, wC, wD = SCENARIO_WEIGHTS["A"], SCENARIO_WEIGHTS["B"], SCENARIO_WEIGHTS["C"], SCENARIO_WEIGHTS["D"]
    l = wA*ta[0] + wB*tb[0] + wC*tc[0] + wD*td[0]
    m = wA*ta[1] + wB*tb[1] + wC*tc[1] + wD*td[1]
    u = wA*ta[2] + wB*tb[2] + wC*tc[2] + wD*td[2]
    return (l, m, u)

def dist_L2(x: tuple[float,float,float], y: tuple[float,float,float]) -> float:
    # Euclidean vertex distance (normalized by 3)
    return math.sqrt(((x[0]-y[0])**2 + (x[1]-y[1])**2 + (x[2]-y[2])**2) / 3.0)

def dist_L1(x: tuple[float,float,float], y: tuple[float,float,float]) -> float:
    return (abs(x[0]-y[0]) + abs(x[1]-y[1]) + abs(x[2]-y[2])) / 3.0

def dist_Linf(x: tuple[float,float,float], y: tuple[float,float,float]) -> float:
    return max(abs(x[0]-y[0]), abs(x[1]-y[1]), abs(x[2]-y[2]))

def cc(d_plus: float, d_minus: float) -> float:
    denom = d_plus + d_minus
    return float(d_minus/denom) if denom != 0 else 0.0

def main() -> None:
    repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    in_path = os.path.join(repo_root, "data", "linguistic_inputs.xlsx")
    out_path = os.path.join(repo_root, "output", "topsis_3dist_results.xlsx")

    xls = pd.ExcelFile(in_path)
    # Read per-dimension sheets (preferred), otherwise try the first sheet.
    preferred = [s for s in ["Economic","Social","Environmental","Technological"] if s in xls.sheet_names]
    if preferred:
        parts = [pd.read_excel(in_path, sheet_name=s) for s in preferred]
        df = pd.concat(parts, ignore_index=True)
    else:
        df = pd.read_excel(in_path, sheet_name=xls.sheet_names[0])

    required = ["Dimension", "Indicator", "Scenario A", "Scenario B", "Scenario C", "Scenario D"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Compute aggregated TFN and 3 distances
    rows = []
    for _, r in df.iterrows():
        dim = r["Dimension"]
        ind = r["Indicator"]
        A, B, C, D = r["Scenario A"], r["Scenario B"], r["Scenario C"], r["Scenario D"]
        x = agg_tfn(A, B, C, D)

        d2p, d2n = dist_L2(x, FPIS), dist_L2(x, FNIS)
        d1p, d1n = dist_L1(x, FPIS), dist_L1(x, FNIS)
        dip, din = dist_Linf(x, FPIS), dist_Linf(x, FNIS)

        rows.append({
            "Dimension": dim,
            "Indicator": ind,
            "Scenario A": _norm_term(A),
            "Scenario B": _norm_term(B),
            "Scenario C": _norm_term(C),
            "Scenario D": _norm_term(D),
            "TFN_l": x[0], "TFN_m": x[1], "TFN_u": x[2],
            "Dplus_L2": d2p, "Dminus_L2": d2n, "CC_L2": cc(d2p, d2n),
            "Dplus_L1": d1p, "Dminus_L1": d1n, "CC_L1": cc(d1p, d1n),
            "Dplus_Linf": dip, "Dminus_Linf": din, "CC_Linf": cc(dip, din),
        })

    out = pd.DataFrame(rows)

    # Ranks within each dimension for each metric
    for metric in ["L2", "L1", "Linf"]:
        out[f"Rank_{metric}"] = out.groupby("Dimension")[f"CC_{metric}"].rank(ascending=False, method="min").astype(int)

    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        out.sort_values(["Dimension", "Rank_L2", "Rank_L1", "Rank_Linf"]).to_excel(w, index=False, sheet_name="TOPSIS_3dist")
        # Also export per-dimension sheets
        for dim, g in out.groupby("Dimension"):
            g2 = g.sort_values(["Rank_L2", "Rank_L1", "Rank_Linf"])
            name = str(dim)[:31] if isinstance(dim, str) else "Dimension"
            g2.to_excel(w, index=False, sheet_name=name)

    print(f"Saved: {out_path}")

if __name__ == "__main__":
    main()
