#!/usr/bin/env python3
"""
Generate Table 3: Overall MRF/E-HR Concordance analysis by indicators.
Outputs a CSV file.

Usage:
  python make_table3_concordance.py --concordance "Concordance_Analysis_Complete.xlsx"
"""

import argparse
import os
import pandas as pd

# Mapping for Concordance Analysis (Wide to Long)
CONCORDANCE_MAP = {
    "HTS_TST":  ("MRF HTS_TX_NEW (Total Tests)",        "DATIM HTS_TX_NEW (Total Tests)"),
    "HTS_POS":  ("MRF HTS_TX_NEW (Total Positives)",    "DATIM HTS_TX_NEW (Total Positives)"),
    "TX_NEW":   ("MRF HTS_TX_NEW (Total Initiations)",  "DATIM HTS_TX_NEW (Total Initiations)"),
    "TX_CURR":  ("MRF TX_CURR",                         "DATIM TX_CURR"),
}

def load_concordance(path: str) -> pd.DataFrame:
    """Load Concordance Analysis Excel and reshape from Wide to Long format."""
    raw = pd.read_excel(path, engine='openpyxl')
    raw.columns = [str(c).strip() for c in raw.columns]
    
    rows = []
    for ind, (mrf_col, ehr_col) in CONCORDANCE_MAP.items():
        if mrf_col in raw.columns and ehr_col in raw.columns:
            tmp = raw.copy()
            tmp["Indicator"] = ind
            tmp["MRF"] = pd.to_numeric(tmp[mrf_col], errors="coerce").fillna(0)
            tmp["EHR"] = pd.to_numeric(tmp[ehr_col], errors="coerce").fillna(0)
            rows.append(tmp[["Indicator", "MRF", "EHR"]])
    
    return pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()

def main():
    parser = argparse.ArgumentParser(description="Generate Table 3 CSV from Concordance Excel")
    parser.add_argument('--concordance', required=True, help='Path to Concordance Analysis Excel')
    parser.add_argument('--outdir', default='output', help='Output directory')
    args = parser.parse_args()

    os.makedirs(args.outdir, exist_ok=True)

    print(f"Loading concordance data from: {args.concordance}")
    df = load_concordance(args.concordance)
    
    if df.empty:
        print("No matching columns found in concordance file.")
        return

    # Aggregate by Indicator
    table3 = df.groupby("Indicator")[["MRF", "EHR"]].sum().reset_index()
    
    # Calculate Concordance
    table3["Concordance"] = table3.apply(
        lambda row: (row["EHR"] / row["MRF"] * 100) if row["MRF"] > 0 else (100.0 if row["EHR"] == 0 else 0.0), 
        axis=1
    )
    
    # Format
    table3["Concordance"] = table3["Concordance"].map('{:.1f}%'.format)
    
    out_csv = os.path.join(args.outdir, 'table3_overall_concordance.csv')
    table3.to_csv(out_csv, index=False)
    print(f"Table 3 CSV saved: {out_csv}")

if __name__ == '__main__':
    main()