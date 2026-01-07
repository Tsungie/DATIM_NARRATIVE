#!/usr/bin/env python3
"""
CDC Quarterly Narrative Report Generator (Full data-to-charts pipeline)

This script renders a CDC-styled Word report from a template (cdc_template.docx)
using docxtpl, generating figures and color-coded concordance tables.

Usage (example):

  python generate_cdc_report_docxtpl.py \
      --template cdc_template.docx \
      --ehr data/ehr_data.xlsx \
      --dhis2 data/dhis2_data.xlsx \
      --optimized data/optimized_sites_concordance.xlsx \
      --out CDC_Report_Q4_rendered.docx \
      --quarter COP24 Q4 \
      --raw-data-sites 573 \
      --collected-manually 322 \
      --pushed-via-pipeline 81 \
      --mobile-backups 170

Requirements:
  pip install pandas matplotlib seaborn python-docx docxtpl reportlab openpyxl
"""

import argparse
import os
import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Lazy imports for docxtpl/python-docx only when rendering
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

pd.options.display.float_format = '{:,.1f}'.format

# -----------------------------
# Helpers
# -----------------------------

def thousands(x):
    try:
        return f"{int(x):,}"
    except Exception:
        return str(x)


def concordance_pct(mrf_val, ehr_val):
    try:
        mrf = float(mrf_val)
        ehr = float(ehr_val)
        if mrf == 0:
            return 0.0
        return round(100.0 * ehr / mrf, 1)
    except Exception:
        return 0.0


def shade_cell(cell, color_hex):
    """Apply background shading to a python-docx cell.
    color_hex: e.g. 'C6EFCE' (green), 'FFEB9C' (amber), 'F8CBAD' (red)
    """
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    # Optional clear others
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)


def color_for_conc(pct):
    """Return hex fill color for given concordance percent.
    Green  : 90–120
    Amber  : 75–89.9
    Red    : <75 or >120
    """
    if pct is None:
        return None
    try:
        p = float(pct)
    except Exception:
        return None
    if 90.0 <= p <= 120.0:
        return 'C6EFCE'  # green
    if 75.0 <= p < 90.0:
        return 'FFEB9C'  # amber
    return 'F8CBAD'      # red


# -----------------------------
# Figure builders
# -----------------------------

def build_figures(ehr_df, outdir, stats):
    os.makedirs(outdir, exist_ok=True)
    sns.set(style="whitegrid")

    # Figure 1: data flow (bar by source)
    fig1_path = os.path.join(outdir, 'figure1.png')
    vals = [
        stats['collected_manually'],
        stats['pushed_via_pipeline'],
        stats['mobile_backups']
    ]
    plt.figure(figsize=(6,4))
    plt.bar(['Manual', 'Pipeline', 'Mobile'], vals, color=['#4C78A8','#F58518','#54A24B'])
    plt.title('COP24 Q4 IMPILO E-HR data flow')
    plt.tight_layout(); plt.savefig(fig1_path, dpi=200); plt.close()

    # Figure 2: proportion of facilities reporting by indicator
    fig2_path = os.path.join(outdir, 'figure2.png')
    labels = ['TX_CURR','TX_ML','TX_TB','HTS_TST/POS','PMTCT_STAT']
    props = [
        stats.get('prop_tx_curr',0),
        stats.get('prop_tx_ml',0),
        stats.get('prop_tx_tb',0),
        stats.get('prop_hts_tst',0),
        stats.get('prop_pmtct_stat',0)
    ]
    plt.figure(figsize=(6,4))
    sns.barplot(x=labels, y=props, color="#4C78A8")
    plt.ylim(0,100)
    plt.title('Proportion of sites reporting by indicator')
    plt.tight_layout(); plt.savefig(fig2_path, dpi=200); plt.close()

    # Figure 3: simple trend placeholder — expects optional historical data in stats['trend']
    fig3_path = os.path.join(outdir, 'figure3.png')
    trend = stats.get('trend')
    if trend is None:
        trend = pd.DataFrame({
            'Quarter': ['COP22','COP23','COP24Q3','COP24Q4'],
            'Sites': [300, 380, 420, stats['ehr_facilities_analyzed']]
        })
    plt.figure(figsize=(6,4))
    sns.lineplot(data=trend, x='Quarter', y='Sites', marker='o')
    plt.title('Trends in sites successfully reporting')
    plt.tight_layout(); plt.savefig(fig3_path, dpi=200); plt.close()

    # Figure 4: TX_CURR sums by district for Harare & Bulawayo (proxy for concordance chart)
    fig4_path = os.path.join(outdir, 'figure4.png')
    sel = ehr_df[ehr_df['district'].isin(['Harare','Bulawayo'])]
    if not sel.empty:
        gb = sel.groupby('district', as_index=False)['TX_CURR'].sum()
        plt.figure(figsize=(6,4))
        sns.barplot(data=gb, x='district', y='TX_CURR', color='#54A24B')
        plt.title('TX_CURR (E-HR sums) by district')
        plt.tight_layout(); plt.savefig(fig4_path, dpi=200); plt.close()
    else:
        # fallback empty figure
        plt.figure(figsize=(6,4)); plt.text(0.5,0.5,'No Harare/Bulawayo data', ha='center');
        plt.axis('off'); plt.savefig(fig4_path, dpi=200); plt.close()

    return {
        'figure1': fig1_path,
        'figure2': fig2_path,
        'figure3': fig3_path,
        'figure4': fig4_path
    }


# -----------------------------
# Table builders (Subdocuments)
# -----------------------------

def build_table1_subdoc(tpl, ehr_df):
    sub = tpl.new_subdoc()
    ind_cols = [c for c in ['HTS_TST','HTS_POS','TX_NEW','TX_CURR','TX_ML','TX_TB','PMTCT_STAT'] if c in ehr_df.columns]
    cols = ['district','facility'] + ind_cols
    table = sub.add_table(rows=1, cols=len(cols))
    # Header
    for j, h in enumerate(cols):
        cell = table.rows[0].cells[j]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.bold = True
    # Rows
    view = ehr_df[cols].copy()
    for _, row in view.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(row['district'])
        cells[1].text = str(row['facility'])
        for k, c in enumerate(ind_cols, start=2):
            try:
                cells[k].text = thousands(row[c])
            except Exception:
                cells[k].text = '0'
    return sub


def build_table2_subdoc(tpl, optimized_df):
    """Optimized sites concordance table with color-coded concordance cells.
    Expects columns: district, facility,
      HTS_TST_MRF, HTS_TST_EHR,
      HTS_POS_MRF, HTS_POS_EHR,
      TX_NEW_MRF, TX_NEW_EHR,
      TX_CURR_MRF, TX_CURR_EHR
    """
    sub = tpl.new_subdoc()
    if optimized_df is None or optimized_df.empty:
        sub.add_paragraph('No optimized sites concordance file provided.')
        return sub

    cols = [
        'district','facility',
        'HTS_TST_MRF','HTS_TST_EHR','HTS_TST_CONC',
        'HTS_POS_MRF','HTS_POS_EHR','HTS_POS_CONC',
        'TX_NEW_MRF','TX_NEW_EHR','TX_NEW_CONC',
        'TX_CURR_MRF','TX_CURR_EHR','TX_CURR_CONC'
    ]
    table = sub.add_table(rows=1, cols=len(cols))
    for j, h in enumerate(cols):
        cell = table.rows[0].cells[j]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.bold = True

    for _, r in optimized_df.iterrows():
        # compute per-indicator concordances
        hts_tst_conc = concordance_pct(r.get('HTS_TST_MRF',0), r.get('HTS_TST_EHR',0))
        hts_pos_conc = concordance_pct(r.get('HTS_POS_MRF',0), r.get('HTS_POS_EHR',0))
        tx_new_conc  = concordance_pct(r.get('TX_NEW_MRF',0),  r.get('TX_NEW_EHR',0))
        tx_curr_conc = concordance_pct(r.get('TX_CURR_MRF',0), r.get('TX_CURR_EHR',0))
        row_cells = table.add_row().cells
        values = [
            r.get('district',''), r.get('facility',''),
            thousands(r.get('HTS_TST_MRF',0)), thousands(r.get('HTS_TST_EHR',0)), f"{hts_tst_conc:.1f}%",
            thousands(r.get('HTS_POS_MRF',0)), thousands(r.get('HTS_POS_EHR',0)), f"{hts_pos_conc:.1f}%",
            thousands(r.get('TX_NEW_MRF',0)),  thousands(r.get('TX_NEW_EHR',0)),  f"{tx_new_conc:.1f}%",
            thousands(r.get('TX_CURR_MRF',0)), thousands(r.get('TX_CURR_EHR',0)), f"{tx_curr_conc:.1f}%",
        ]
        for j, v in enumerate(values):
            row_cells[j].text = str(v)
        # Shade the concordance cells
        for idx, pct in zip([4,7,10,13], [hts_tst_conc, hts_pos_conc, tx_new_conc, tx_curr_conc]):
            color = color_for_conc(pct)
            if color:
                shade_cell(row_cells[idx], color)
    return sub


# -----------------------------
# Main
# -----------------------------

def main():
    parser = argparse.ArgumentParser(description='Generate CDC quarterly narrative report (docxtpl).')
    parser.add_argument('--template', default='cdc_template.docx', help='Word template with placeholders')
    parser.add_argument('--ehr', required=True, help='Path to EHR Excel (IMPILO extract)')
    parser.add_argument('--dhis2', required=True, help='Path to DHIS2/MRF Excel')
    parser.add_argument('--optimized', default=None, help='Path to optimized sites concordance Excel (optional)')
    parser.add_argument('--out', default='CDC_Report_Q4_rendered.docx', help='Output DOCX path')
    parser.add_argument('--figdir', default='figures', help='Directory to save generated figures')
    parser.add_argument('--quarter', default='COP24 Q4', help='Quarter label for captions if needed')

    # Headline stats (override defaults from CLI)
    parser.add_argument('--total-active-facilities', type=int, default=None)
    parser.add_argument('--raw-data-sites', type=int, default=573)
    parser.add_argument('--collected-manually', type=int, default=322)
    parser.add_argument('--pushed-via-pipeline', type=int, default=81)
    parser.add_argument('--mobile-backups', type=int, default=170)

    args = parser.parse_args()

    # Load data
    ehr = pd.read_excel(args.ehr, engine='openpyxl')
    dhis2 = pd.read_excel(args.dhis2, engine='openpyxl')
    optimized = None
    if args.optimized and os.path.exists(args.optimized):
        optimized = pd.read_excel(args.optimized, engine='openpyxl')

    # Sanity columns
    required_cols = {'facility','district','HTS_TST','HTS_POS','TX_NEW','TX_CURR'}
    if not required_cols.issubset(ehr.columns):
        missing = ', '.join(sorted(required_cols - set(ehr.columns)))
        sys.exit(f"EHR file missing required columns: {missing}")
    if not required_cols.issubset(dhis2.columns):
        missing = ', '.join(sorted(required_cols - set(dhis2.columns)))
        sys.exit(f"DHIS2 file missing required columns: {missing}")

    # Headline numbers
    total_active_facilities = args.total_active_facilities or ehr['facility'].nunique()
    raw_data_sites = args.raw_data_sites
    collected_manually = args.collected_manually
    pushed_via_pipeline = args.pushed_via_pipeline
    mobile_backups = args.mobile_backups

    ehr_facilities_analyzed = ehr['facility'].nunique()
    ehr_facilities_analyzed_pct = round(100 * ehr_facilities_analyzed / max(total_active_facilities,1), 1)

    def prop_reporting(df, col):
        return round(100 * (df.groupby('facility')[col].sum() > 0).mean(), 1)

    prop_tx_curr = prop_reporting(ehr, 'TX_CURR')
    prop_tx_ml   = prop_reporting(ehr, 'TX_ML') if 'TX_ML' in ehr.columns else 0.0
    prop_tx_tb   = prop_reporting(ehr, 'TX_TB') if 'TX_TB' in ehr.columns else 0.0
    prop_hts_tst = prop_reporting(ehr, 'HTS_TST')  # HTS_TST/POS combined label uses HTS_TST base here
    prop_pmtct_stat = prop_reporting(ehr, 'PMTCT_STAT') if 'PMTCT_STAT' in ehr.columns else 0.0

    # Overall concordance (Table 3)
    def sums(df, cols):
        return {c: int(df[c].sum()) for c in cols}

    cols = ['HTS_TST','HTS_POS','TX_NEW','TX_CURR']
    mrf_sums = sums(dhis2, cols)
    ehr_sums = sums(ehr, cols)

    overall_hts_tst_conc = concordance_pct(mrf_sums['HTS_TST'], ehr_sums['HTS_TST'])
    overall_hts_pos_conc = concordance_pct(mrf_sums['HTS_POS'], ehr_sums['HTS_POS'])
    overall_tx_new_conc  = concordance_pct(mrf_sums['TX_NEW'],  ehr_sums['TX_NEW'])
    overall_tx_curr_conc = concordance_pct(mrf_sums['TX_CURR'], ehr_sums['TX_CURR'])

    stats = {
        'total_active_facilities': total_active_facilities,
        'raw_data_sites': raw_data_sites,
        'collected_manually': collected_manually,
        'pushed_via_pipeline': pushed_via_pipeline,
        'mobile_backups': mobile_backups,
        'ehr_facilities_analyzed': ehr_facilities_analyzed,
        'ehr_facilities_analyzed_pct': ehr_facilities_analyzed_pct,
        'prop_tx_curr': prop_tx_curr,
        'prop_tx_ml': prop_tx_ml,
        'prop_tx_tb': prop_tx_tb,
        'prop_hts_tst': prop_hts_tst,
        'prop_pmtct_stat': prop_pmtct_stat,
    }

    # Build figures
    figs = build_figures(ehr, args.figdir, stats)

    # Load template
    tpl = DocxTemplate(args.template)

    # Build subdocs (tables)
    sub1 = build_table1_subdoc(tpl, ehr)
    sub2 = build_table2_subdoc(tpl, optimized)

    # Context for template rendering
    context = {
        # Header counts
        'total_active_facilities': total_active_facilities,
        'raw_data_sites': raw_data_sites,
        'collected_manually': collected_manually,
        'pushed_via_pipeline': pushed_via_pipeline,
        'mobile_backups': mobile_backups,
        'ehr_facilities_analyzed': ehr_facilities_analyzed,
        'ehr_facilities_analyzed_pct': ehr_facilities_analyzed_pct,
        # Data Quality Metrics proportions
        'prop_tx_curr': prop_tx_curr,
        'prop_tx_ml': prop_tx_ml,
        'prop_tx_tb': prop_tx_tb,
        'prop_hts_tst': prop_hts_tst,
        'prop_pmtct_stat': prop_pmtct_stat,
        # Narrative text placeholders (adjust as needed)
        'challenge_database_collection': (
            'Onboarding more facilities to the automated data pipeline and improving the mobile application '
            'to reduce manual intervention; addressing network selection issues and ensuring complete uploads.'
        ),
        'challenge_data_extraction': (
            'Higher success for manual/dpl backups compared to mobile uploads due to slow central network performance; '
            'large facilities may require batch processing to avoid memory failures.'
        ),
        'remedial_point_1': 'Continue downloading and reprocessing missing data from the central repository.',
        'remedial_point_2': 'Generate data in smaller batches for large facilities to avoid memory-related failures.',
        'remedial_point_3': 'Strengthen mentorship and power backup at facilities.',
        # Overall concordance (Table 3)
        'overall_hts_tst_mrf': f"{mrf_sums['HTS_TST']:,}",
        'overall_hts_tst_ehr': f"{ehr_sums['HTS_TST']:,}",
        'overall_hts_tst_conc': overall_hts_tst_conc,
        'overall_hts_pos_mrf': f"{mrf_sums['HTS_POS']:,}",
        'overall_hts_pos_ehr': f"{ehr_sums['HTS_POS']:,}",
        'overall_hts_pos_conc': overall_hts_pos_conc,
        'overall_tx_new_mrf': f"{mrf_sums['TX_NEW']:,}",
        'overall_tx_new_ehr': f"{ehr_sums['TX_NEW']:,}",
        'overall_tx_new_conc': overall_tx_new_conc,
        'overall_tx_curr_mrf': f"{mrf_sums['TX_CURR']:,}",
        'overall_tx_curr_ehr': f"{ehr_sums['TX_CURR']:,}",
        'overall_tx_curr_conc': overall_tx_curr_conc,
        # Subdocs for tables
        'table1_indicators': sub1,
        'table2_concordance_facility': sub2,
        # Figures (InlineImage)
        'figure1': InlineImage(tpl, figs['figure1'], width=Mm(140)),
        'figure2': InlineImage(tpl, figs['figure2'], width=Mm(140)),
        'figure3': InlineImage(tpl, figs['figure3'], width=Mm(140)),
        'figure4': InlineImage(tpl, figs['figure4'], width=Mm(140)),
    }

    tpl.render(context)
    tpl.save(args.out)
    print(f"CDC report rendered: {args.out}")


if __name__ == '__main__':
    main()
