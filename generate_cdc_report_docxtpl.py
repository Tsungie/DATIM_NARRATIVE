#!/usr/bin/env python3
"""
Generate CDC Figure 2 and Table 1 from Consistency Report (COP24 Q4)

This script:
  - Loads a Consistency_Report Excel (Indicator, Number of Reporting Units)
  - Computes reporting completeness for a fixed set of 21 indicators vs targets
  - Produces Figure 2 (PNG) and a Table 1 subdocument
  - Renders into a Word template (default: cdc_template.docx)

Usage:
  python generate_cdc_report_consistency.py \
    --consistency Consistency_Report_COP2024_Q4.xlsx \
    --template "cdc_template.docx" \
    --out output/CDC_Report_COP24_Q4_rendered.docx \
    --period COP24 Q4

Dependencies (install in your venv):
  pip install pandas matplotlib seaborn docxtpl python-docx openpyxl docxcompose
"""

import argparse
import os
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# ---------------
# Targets (fixed 21 indicators)
# ---------------
TARGETS = {
    'TX_CURR': 823,
    'TX_TB': 823,
    'TX_ML': 823,
    'HTS_TST and HTS_POS': 823,
    'PMTCT_STAT': 761,
    'PMTCT_ART': 659,
    'TB_PREV': 657,
    'TX_PVLS': 823,
    'HTS_INDEX': 823,
    'TX_NEW': 823,
    'TB_ART': 657,
    'PREP_NEW': 628,
    'PMTCT_FO': 661,
    'HTS_SELF': 823,
    'TB_STAT': 657,
    'PMTCT_EID': 655,
    'CXCA_SCRN': 628,
    'CXCA_TX': 473,
    'PREP_CT': 628,
    'TX_RTT': 823,
    'PMTCT_HEI': 661,
}

# ---------------
# Mapping rules: Consistency Report indicator names -> canonical
# Prefer denominators when available
# ---------------
MAP_RULES = {
    'TX_CURR': ['TX_CURR TA'],
    'TX_TB': ['TX_TB(DENOM) TA', 'TX_TB TA'],
    'TX_ML': ['TX_ML TA'],
    'HTS_TST': ['HTS_TST TA'],
    'HTS_POS': ['HTS_POS TA'],  # may be absent
    'PMTCT_STAT': ['PMTCT_STAT(DENOM) TA', 'PMTCT_STAT TA'],
    'PMTCT_ART': ['PMTCT_ART TA'],
    'TB_PREV': ['TB_PREV(DENOM) TA', 'TB_PREV TA'],
    'TX_PVLS': ['TX_PVLS(DENOM) TA', 'TX_PVLS TA'],
    'HTS_INDEX': ['HTS_INDEX TA'],
    'TX_NEW': ['TX_NEW TA'],
    'TB_ART': ['TB_ART TA'],
    'PREP_NEW': ['PrEP_NEW TA', 'PREP_NEW TA'],
    'PMTCT_FO': ['PMTCT_FO TA'],
    'HTS_SELF': ['HTS_SELF TA'],
    'TB_STAT': ['TB_STAT(DENOM) TA', 'TB_STAT TA'],
    'PMTCT_EID': ['PMTCT_EID TA'],
    'CXCA_SCRN': ['CXCA_SCRN TA'],
    'CXCA_TX': ['CXCA_TX TA'],
    'PREP_CT': ['PrEP_CT TA', 'PREP_CT TA'],
    'TX_RTT': ['TX_RTT TA'],
    'PMTCT_HEI': ['PMTCT_HEI TA'],
}

INDICATOR_ORDER = [
    'TX_CURR', 'TX_TB', 'TX_ML', 'HTS_TST and HTS_POS', 'PMTCT_STAT', 'PMTCT_ART',
    'TB_PREV', 'TX_PVLS', 'HTS_INDEX', 'TX_NEW', 'TB_ART', 'PREP_NEW', 'PMTCT_FO',
    'HTS_SELF', 'TB_STAT', 'PMTCT_EID', 'CXCA_SCRN', 'CXCA_TX', 'PREP_CT', 'TX_RTT', 'PMTCT_HEI'
]


def load_consistency(path: str) -> pd.DataFrame:
    """Load Consistency Report and return DataFrame with columns ['Indicator','ReportingUnits']."""
    df = pd.read_excel(path, engine='openpyxl', sheet_name=0)
    # Heuristic: find the two required columns
    ind_col = None
    num_col = None
    for c in df.columns:
        cname = str(c).strip().lower()
        if 'indicator' in cname:
            ind_col = c
        if ('number' in cname or 'no.' in cname) and 'report' in cname:
            num_col = c
    if ind_col is None or num_col is None:
        # Fallback to first two columns
        ind_col = df.columns[0]
        num_col = df.columns[1]
    out = pd.DataFrame({
        'Indicator': df[ind_col].astype(str).str.strip(),
        'ReportingUnits': pd.to_numeric(df[num_col], errors='coerce').fillna(0).astype(int)
    })
    return out


def pick_reporting(cons_df: pd.DataFrame, patterns: list[str]) -> int:
    """Pick reporting units by trying patterns in order (first hit wins) and returning the value."""
    for pat in patterns:
        row = cons_df[cons_df['Indicator'] == pat]
        if not row.empty:
            return int(row.iloc[0]['ReportingUnits'])
    return 0


def build_21_indicator_table(cons_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    # First compute individual pieces (including HTS_TST & HTS_POS)
    values = {}
    for key, pats in MAP_RULES.items():
        values[key] = pick_reporting(cons_df, pats)
    # Combined HTS_TST and HTS_POS: choose max if both present, else whichever is present
    hts_combined = max(values.get('HTS_TST', 0), values.get('HTS_POS', 0))

    # Compose the 21-indicator table
    for ind in INDICATOR_ORDER:
        if ind == 'HTS_TST and HTS_POS':
            reporting = hts_combined
        else:
            reporting = values.get(ind, 0)
        tgt = TARGETS[ind]
        comp = round(100 * reporting / tgt, 1) if tgt else 0.0
        rows.append({
            'Indicator': ind,
            'Number_of_facilities_reporting': reporting,
            'Target': tgt,
            'Reporting_Completeness_%': comp,
        })
    return pd.DataFrame(rows)


def build_figure2(table_df: pd.DataFrame, out_png: str, period_label: str):
    sns.set(style='whitegrid')
    plt.figure(figsize=(12,8))
    plot_df = table_df.sort_values('Reporting_Completeness_%', ascending=False)
    ax = sns.barplot(x='Indicator', y='Reporting_Completeness_%', data=plot_df, color='#4C78A8')
    plt.ylabel('Reporting Completeness (%)')
    plt.xlabel('Indicator')
    plt.ylim(0,100)
    plt.title(f'Figure 2: Proportion of sites reporting individual indicators in IMPILO E-HR {period_label}')
    plt.xticks(rotation=45, ha='right')
    # annotate bars with % only
    for i, p in enumerate(ax.patches):
        pct = plot_df.iloc[i]['Reporting_Completeness_%']
        ax.annotate(f"{pct}%", (p.get_x()+p.get_width()/2, p.get_height()),
                    ha='center', va='bottom', fontsize=9)
    plt.tight_layout(); plt.savefig(out_png, dpi=200); plt.close()


def build_table1_subdoc(tpl: DocxTemplate, table_df: pd.DataFrame):
    sub = tpl.new_subdoc()
    # Create table with 4 columns
    tbl = sub.add_table(rows=1, cols=4)
    headers = ['Indicator', 'Number of facilities reporting', 'Target', 'Reporting Completeness']
    for j, h in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.bold = True
    # Add rows in the predefined order
    for _, r in table_df.iterrows():
        row = tbl.add_row().cells
        row[0].text = str(r['Indicator'])
        row[1].text = str(int(r['Number_of_facilities_reporting']))
        row[2].text = str(int(r['Target']))
        row[3].text = f"{r['Reporting_Completeness_%']}%"
    return sub


def main():
    parser = argparse.ArgumentParser(description='Generate CDC Report Figure 2 & Table 1 from Consistency Report.')
    parser.add_argument('--consistency', required=True, help='Path to Consistency_Report Excel')
    parser.add_argument('--template', default='cdc_template.docx', help='Path to CDC report template DOCX (default: cdc_template.docx)')
    parser.add_argument('--out', default=None, help='Output DOCX path (default: output/CDC_Report_<period>_<date>.docx)')
    parser.add_argument('--period', default='COP24 Q4', help='Label used in figure title')
    parser.add_argument('--outdir', default='output', help='Directory for generated assets')
    args = parser.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    stamp = datetime.now().strftime('%Y-%m-%d')

    # Load consistency report
    cons_df = load_consistency(args.consistency)

    # Build 21-indicator table
    table_df = build_21_indicator_table(cons_df)

    # Save CSV
    csv_path = os.path.join(args.outdir, f'figure2_table1_{args.period.replace(" ", "")}_{stamp}.csv')
    table_df.to_csv(csv_path, index=False)

    # Build Figure 2
    fig_path = os.path.join(args.outdir, f'figure2_reporting_completeness_{args.period.replace(" ", "")}_{stamp}.png')
    build_figure2(table_df, fig_path, args.period)

    # Load template and render
    tpl = DocxTemplate(args.template)
    sub_table1 = build_table1_subdoc(tpl, table_df)
    context = {
        'figure2': InlineImage(tpl, fig_path, width=Mm(140)),
        'table1_indicators': sub_table1,
    }
    tpl.render(context)

    out_docx = args.out or os.path.join(args.outdir, f'CDC_Report_{args.period.replace(" ", "")}_{stamp}.docx')
    tpl.save(out_docx)
    print('Figure 2 PNG:', fig_path)
    print('Table 1 CSV:', csv_path)
    print('Rendered report:', out_docx)

if __name__ == '__main__':
    main()
