
#!/usr/bin/env python3
"""
Create Figure 2: Proportion of sites reporting individual indicators
using a DATIM Excel and indicator-specific Target facility counts.

Outputs (in output/ with today’s date):
- CSV: figure2_reporting_completeness_<period>_<YYYY-MM-DD>.csv
- PNG: figure2_reporting_completeness_<period>_<YYYY-MM-DD>.png

Usage example:
  python make_figure2_from_datim.py \
      --datim "DATIM_Report_COP24 Q4.xlsx" \
      --period 2025Q4 \
      --outdir output
"""

import os
import argparse
from datetime import datetime
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # non-interactive backend
import matplotlib.pyplot as plt
import seaborn as sns

def load_datim(datim_path: str) -> pd.DataFrame:
    usecols = ["Period", "SiteProvince", "SiteDistrict", "SiteName",
               "Indicator", "DATIM_Indicator", "Value"]
    df = pd.read_excel(datim_path, engine="openpyxl", usecols=usecols)
    df.columns = [c.strip().replace(" ", "_") for c in df.columns]
    return df

def compute_facilities_reporting(df: pd.DataFrame, period: str) -> pd.DataFrame:
    period_df = df[df["Period"].astype(str) == period].copy()
    per_fac = (period_df.groupby(["Indicator", "SiteName"], as_index=False)["Value"].sum())
    per_fac["reported"] = per_fac["Value"] > 0

    fac_counts = (per_fac[per_fac["reported"]]
                  .groupby("Indicator")["SiteName"]
                  .nunique()
                  .rename("Number_of_facilities_reporting")
                  .reset_index())

    # Combined HTS_TST and HTS_POS
    hts_tst_sites = set(per_fac[(per_fac["Indicator"] == "HTS_TST") & (per_fac["reported"])]["SiteName"])
    hts_pos_sites = set(per_fac[(per_fac["Indicator"] == "HTS_POS") & (per_fac["reported"])]["SiteName"])
    combined_hts_sites = hts_tst_sites.union(hts_pos_sites)
    combined_row = pd.DataFrame({
        "Indicator": ["HTS_TST and HTS_POS"],
        "Number_of_facilities_reporting": [len(combined_hts_sites)]
    })
    fac_counts = pd.concat([fac_counts, combined_row], ignore_index=True)

    return fac_counts

def build_completeness_table(fac_counts: pd.DataFrame, targets: dict) -> pd.DataFrame:
    rows = []
    for ind, tgt in targets.items():
        n_reporting = int(fac_counts.loc[fac_counts["Indicator"] == ind,
                                         "Number_of_facilities_reporting"].sum())
        completeness = round(100 * n_reporting / tgt, 1) if tgt else 0.0
        rows.append({
            "Indicator": ind,
            "Number_of_facilities_reporting": n_reporting,

            "Reporting_Completeness_%": completeness,
        })
    return pd.DataFrame(rows)


def plot_figure2(df: pd.DataFrame, period: str, out_png: str) -> None:
    import matplotlib.ticker as mtick
    sns.set(style="whitegrid")
    plt.figure(figsize=(12, 8))

    plot_df = df.sort_values("Reporting_Completeness_%", ascending=False)

    ax = sns.barplot(
        x="Indicator",
        y="Reporting_Completeness_%",
        data=plot_df,
        color="#4C78A8",
        edgecolor="black"
    )

    # Axis labels & title
    ax.set_ylabel("Reporting Completeness (%)")
    ax.set_xlabel("Indicator")
    ax.set_title(f"Proportion of sites reporting individual indicators in IMPILO E-HR {period}")
    ax.set_ylim(0, 100)
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))  # show 0–100% as percent ticks

    # Rotate x-labels
    plt.xticks(rotation=45, ha="right")

    # Grid
    ax.grid(True, axis="y", linestyle="--", alpha=0.3)

    # Percent-only labels (no n/target)
    for i, p in enumerate(ax.patches):
        pct = plot_df.iloc[i]["Reporting_Completeness_%"]
        # Place label slightly above the bar
        ax.annotate(f"{pct}%", (p.get_x() + p.get_width() / 2, p.get_height()),
                    ha="center", va="bottom", fontsize=9)

    plt.tight_layout()
    plt.savefig(out_png, dpi=200)
    plt.close()

def main():
    ap = argparse.ArgumentParser(description="Generate Figure 2 from DATIM Excel and targets.")
    ap.add_argument("--datim", required=True, help="Path to DATIM Excel (facility-level)")
    ap.add_argument("--period", required=False, default=None,
                    help="Period to filter (e.g., 2025Q4). If omitted, last period in file is used.")
    ap.add_argument("--targets", required=False, default=None,
                    help="Optional CSV mapping (Indicator,Target). If omitted, built-in targets are used.")
    ap.add_argument("--outdir", default="output", help="Output directory")
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d")

    df = load_datim(args.datim)
    period = args.period or sorted(df["Period"].astype(str).unique())[-1]

    fac_counts = compute_facilities_reporting(df, period)
    targets = load_targets(args.targets)
    completeness_df = build_completeness_table(fac_counts, targets)

    out_csv = os.path.join(args.outdir, f"figure2_reporting_completeness_{period}_{stamp}.csv")
    completeness_df.to_csv(out_csv, index=False)

    out_png = os.path.join(args.outdir, f"figure2_reporting_completeness_{period}_{stamp}.png")
    plot_figure2(completeness_df, period, out_png)

    print("Period:", period)
    print("Saved CSV:", out_csv)
    print("Saved PNG:", out_png)

if __name__ == "__main__":
    main()
