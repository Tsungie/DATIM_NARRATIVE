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
import matplotlib
matplotlib.use('Agg')  
import matplotlib.pyplot as plt
import seaborn as sns

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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

# Mapping for Concordance Analysis (Wide to Long)
CONCORDANCE_MAP = {
    "HTS_TST":  ("MRF HTS_TX_NEW (Total Tests)",        "DATIM HTS_TX_NEW (Total Tests)"),
    "HTS_POS":  ("MRF HTS_TX_NEW (Total Positives)",    "DATIM HTS_TX_NEW (Total Positives)"),
    "TX_NEW":   ("MRF HTS_TX_NEW (Total Initiations)",  "DATIM HTS_TX_NEW (Total Initiations)"),
    "TX_CURR":  ("MRF TX_CURR",                         "DATIM TX_CURR"),
}


def load_consistency(path: str) -> pd.DataFrame:
    """Load Consistency Report and return DataFrame with columns ['Indicator','ReportingUnits']."""
    df = pd.read_excel(path, engine='openpyxl', sheet_name=0)
    
    # Ensure unique columns to prevent DataFrame selection errors
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]

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


def concordance_value(mrf, ehr):
    """Calculate concordance percentage."""
    if mrf == 0 and ehr == 0: return 100.0
    if mrf == 0 and ehr > 0: return 0.0
    return max(0.0, min(100.0, (1 - abs(ehr - mrf) / mrf) * 100))


def load_concordance(path: str) -> pd.DataFrame:
    """Load Concordance Analysis Excel and reshape from Wide to Long format."""
    raw = pd.read_excel(path, engine='openpyxl')
    raw.columns = [str(c).strip() for c in raw.columns]
    
    rows = []
    for ind, (mrf_col, ehr_col) in CONCORDANCE_MAP.items():
        # Skip if columns missing
        if mrf_col not in raw.columns or ehr_col not in raw.columns:
            continue
            
        tmp = raw[["Facility", "Province", "District", mrf_col, ehr_col]].copy()
        tmp["Indicator"] = ind
        tmp["MRF"] = pd.to_numeric(tmp[mrf_col], errors="coerce").fillna(0)
        tmp["EHR"] = pd.to_numeric(tmp[ehr_col], errors="coerce").fillna(0)
        tmp["Concordance_%"] = tmp.apply(lambda r: concordance_value(r["MRF"], r["EHR"]), axis=1)
        
        rows.append(tmp[["Facility", "Province", "District", "Indicator", "MRF", "EHR", "Concordance_%"]])
    
    return pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()


def shade_cell(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def get_concordance_color(val):
    if val >= 95: return 'C6EFCE' # Green
    if val >= 90: return 'FFEB9C' # Amber
    return 'F8CBAD' # Red


def compute_table3_context(df: pd.DataFrame) -> dict:
    """Compute Overall Data Concordance for Table 3."""
    context = {}
    required_cols = {'Indicator', 'MRF', 'EHR'}
    if not required_cols.issubset(df.columns):
        print(f"Warning: Concordance file missing columns for Table 3. Found: {df.columns}")
        return context

    # Indicators for Table 3
    indicators = ['HTS_TST', 'HTS_POS', 'TX_NEW', 'TX_CURR']
    
    for ind in indicators:
        # Filter by indicator (case insensitive)
        sub = df[df['Indicator'].astype(str).str.upper() == ind]
        
        mrf_sum = sub['MRF'].sum() if not sub.empty else 0
        ehr_sum = sub['EHR'].sum() if not sub.empty else 0
        
        # Calculate concordance
        conc = round((ehr_sum / mrf_sum * 100), 1) if mrf_sum > 0 else 0.0
        
        # Populate context keys (lowercase keys as per template)
        key_base = f"overall_{ind.lower()}"
        context[f"{key_base}_mrf"] = f"{int(mrf_sum):,}"
        context[f"{key_base}_ehr"] = f"{int(ehr_sum):,}"
        context[f"{key_base}_conc"] = str(conc)

    return context


def build_figure4(df: pd.DataFrame, out_png: str):
    """
    Generate TX_CURR Concordance Combo Chart.
    - Left Axis: Grouped Bars (Volume)
    - Right Axis: Hidden (Red Diamonds float at correct height)
    - Legend: Outside on the right, includes Concordance.
    """
    import matplotlib.patches as mpatches
    from matplotlib.lines import Line2D

    required_cols = {'Indicator', 'Province', 'MRF', 'EHR'}
    if not required_cols.issubset(df.columns):
        print(f"Warning: Concordance file missing columns for Figure 4. Found: {df.columns}")
        return

    # 1. Filter and Normalize
    target_provinces = ['Harare', 'Bulawayo']
    
    mask_prov = df['Province'].astype(str).str.strip().apply(
        lambda x: any(p.lower() in x.lower() for p in target_provinces)
    )
    mask_ind = df['Indicator'].astype(str).str.upper() == 'TX_CURR'
    subset = df[mask_prov & mask_ind].copy()
    
    if subset.empty:
        print("Warning: No data found for Figure 4.")
        return

    def normalize_name(p):
        p_str = str(p).lower()
        if 'harare' in p_str: return 'Harare'
        if 'bulawayo' in p_str: return 'Bulawayo'
        return p
    subset['Province'] = subset['Province'].apply(normalize_name)

    # 2. Aggregate
    grouped = subset.groupby('Province')[['MRF', 'EHR']].sum().reset_index()
    
    # Create Total Row
    total_row = pd.DataFrame({
        'Province': ['Total'],
        'MRF': [grouped['MRF'].sum()],
        'EHR': [grouped['EHR'].sum()]
    })
    ordered_provinces = pd.concat([total_row, grouped], ignore_index=True)

    # 3. Prepare Data
    plot_df = ordered_provinces.melt(id_vars='Province', 
                                     value_vars=['MRF', 'EHR'], 
                                     var_name='Source', 
                                     value_name='Count')
    
    # Calculate Concordance %
    ordered_provinces['Concordance'] = ordered_provinces.apply(
        lambda r: (r['EHR'] / r['MRF'] * 100) if r['MRF'] > 0 else 0, axis=1
    )

    x_order = ['Total'] + [p for p in target_provinces if p in grouped['Province'].unique()]

    # 4. Plotting
    sns.set(style='white')
    # Increase figure width to accommodate the external legend
    fig, ax1 = plt.subplots(figsize=(11, 6))

    # --- LEFT AXIS: BARS ---
    custom_colors = {'MRF': '#104E8B', 'EHR': '#00FF00'} # Royal Blue & Bright Green
    
    # We turn off the automatic legend here because we will build a custom one
    sns.barplot(
        x='Province', 
        y='Count', 
        hue='Source', 
        data=plot_df, 
        order=x_order,
        palette=custom_colors, 
        edgecolor="black",
        ax=ax1
    )
    
    # Remove default legend
    if ax1.get_legend():
        ax1.get_legend().remove()
    
    ax1.set_ylabel('Number of clients on ART', fontsize=12)
    ax1.set_xlabel('Province', fontsize=12)
    ax1.set_title('TX_CURR Concordance: MRF vs EHR', fontsize=14, pad=20)
    
    # Add Bar Labels
    for container in ax1.containers:
        ax1.bar_label(container, fmt='{:,.0f}', padding=3, fontsize=10)

    # --- RIGHT AXIS: DIAMONDS ---
    ax2 = ax1.twinx()
    ax2.set_ylim(0, 115) # Keep scale for positioning
    
    # Extract y-values for the diamonds
    y_values = []
    for region in x_order:
        val = ordered_provinces.loc[ordered_provinces['Province'] == region, 'Concordance'].values[0]
        y_values.append(val)
    
    x_coords = range(len(x_order))
    
    # Plot Diamonds
    ax2.plot(x_coords, y_values, color='#D62728', marker='D', markersize=8, linestyle='None')

    # Add Percentage Labels above Diamonds
    for x, y in zip(x_coords, y_values):
        ax2.text(x, y + 4, f"{y:.1f}%", color='#D62728', ha='center', fontweight='bold', fontsize=11)

    # --- CLEANUP & LEGEND ---
    
    # Hide the secondary axis visuals (ticks, numbers, spine)
    ax2.set_yticks([])
    ax2.set_ylabel('')
    ax2.spines['right'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)

    # Build Custom Legend (Key)
    legend_elements = [
        mpatches.Patch(facecolor='#104E8B', edgecolor='black', label='MRF'),
        mpatches.Patch(facecolor='#00FF00', edgecolor='black', label='EHR'),
        Line2D([0], [0], color='#D62728', marker='D', linestyle='None', markersize=8, label='Concordance')
    ]

    # Place legend outside to the right
    ax1.legend(handles=legend_elements, title='Key', bbox_to_anchor=(1.02, 1), loc='upper left')

    ax1.grid(True, axis='y', alpha=0.3)

    plt.tight_layout()
    plt.savefig(out_png, dpi=200, bbox_inches='tight') # bbox_inches ensures legend isn't cut off
    plt.close()
def aggregate_table3_data(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate raw concordance data for Table 3."""
    indicators = ['HTS_TST', 'HTS_POS', 'TX_NEW', 'TX_CURR']
    data = []
    for ind in indicators:
        # Filter by indicator (case insensitive)
        sub_df = df[df['Indicator'].astype(str).str.upper() == ind]
        
        mrf_sum = sub_df['MRF'].sum() if not sub_df.empty else 0
        ehr_sum = sub_df['EHR'].sum() if not sub_df.empty else 0
        
        # Calculate concordance
        conc = (ehr_sum / mrf_sum * 100) if mrf_sum > 0 else (100.0 if ehr_sum == 0 else 0.0)
        
        data.append({
            "Indicator": ind,
            "MRF/DHIS2": int(mrf_sum),
            "E-HR": int(ehr_sum),
            "Concordance": conc
        })
    return pd.DataFrame(data)


def build_overall_table_subdoc(tpl: DocxTemplate, agg_df: pd.DataFrame):
    """Build Table 3 (Overall Concordance) as a subdocument."""
    sub = tpl.new_subdoc()
    if agg_df.empty:
        sub.add_paragraph("No data available")
        return sub

    # Create table
    tbl = sub.add_table(rows=1, cols=4)
    try:
        tbl.style = 'Table Grid'
    except KeyError:
        pass
        
    headers = ["Indicator", "MRF/DHIS2", "E-HR", "Concordance"]
    hdr_cells = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        for p in hdr_cells[i].paragraphs:
            for run in p.runs:
                run.bold = True

    for _, row in agg_df.iterrows():
        row_cells = tbl.add_row().cells
        row_cells[0].text = str(row["Indicator"])
        row_cells[1].text = f"{int(row['MRF/DHIS2']):,}"
        row_cells[2].text = f"{int(row['E-HR']):,}"
        conc = row["Concordance"]
        row_cells[3].text = f"{conc:.1f}%"
        
        # Color coding
        color = get_concordance_color(conc)
        if color:
            shade_cell(row_cells[3], color)
            
    return sub


def build_table2_subdoc(tpl: DocxTemplate, df: pd.DataFrame):
    sub = tpl.new_subdoc()
    if df.empty:
        sub.add_paragraph("No data available")
        return sub

    # Add table
    cols = 7  # Fixed columns: Facility, Province, District, Indicator, MRF, EHR, Concordance
    tbl = sub.add_table(rows=1, cols=cols)
    try:
        tbl.style = 'Table Grid'
    except KeyError:
        pass

    # Header
    headers = ["Facility", "Province", "District", "Indicator", "MRF", "EHR", "Concordance_%"]
    hdr_cells = tbl.rows[0].cells
    for i, col in enumerate(headers):
        hdr_cells[i].text = str(col)
        for p in hdr_cells[i].paragraphs:
            for run in p.runs:
                run.bold = True

    # Rows
    for _, row_series in df.iterrows():
        # Ensure we grab data in correct order
        row_cells = tbl.add_row().cells
        
        # Map row_series data to fixed column order
        row_data = [row_series.get(h, "") for h in headers]
        
        for i, val in enumerate(row_data):
            txt = "" if pd.isna(val) else str(val)
            if isinstance(val, float) and val.is_integer():
                txt = str(int(val))
            elif isinstance(val, float) and headers[i] == "Concordance_%":
                txt = f"{val:.1f}%"
                
            row_cells[i].text = txt
            
            # Color coding for Concordance column
            if headers[i] == "Concordance_%":
                try:
                    val_float = float(val)
                    color = get_concordance_color(val_float)
                    if color: shade_cell(row_cells[i], color)
                    row_cells[i].text = f"{val_float:.1f}%"
                except (ValueError, TypeError):
                    pass
    return sub


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
    parser.add_argument('--concordance', help='Path to Concordance_Analysis_Complete Excel for Table 2')
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

    # Save CSV figure2_table1
    csv_path = os.path.join(args.outdir, f'figure2_table1_{args.period.replace(" ", "")}_{stamp}.csv')
    table_df.to_csv(csv_path, index=False)

    # Build Figure 2
    fig_path = os.path.join(args.outdir, f'figure2_reporting_completeness_{args.period.replace(" ", "")}_{stamp}.png')
    build_figure2(table_df, fig_path, args.period)

    # Load template and render
    tpl = DocxTemplate(args.template)
    sub_table1 = build_table1_subdoc(tpl, table_df)
    
    sub_table2 = ""
    table3_context = {}
    fig4_path = ""
    
    if args.concordance:
        conc_df = load_concordance(args.concordance)
        
        # Save Table 2 CSV (merging functionality from make_table2_concordance.py)
        t2_csv = os.path.join(args.outdir, f'table2_concordance_{args.period.replace(" ", "")}_{stamp}.csv')
        conc_df.to_csv(t2_csv, index=False)
        print('Table 2 CSV:', t2_csv)

        # Generate Table 3 Data & CSV
        table3_df = aggregate_table3_data(conc_df)
        t3_csv = os.path.join(args.outdir, f'table3_overall_concordance_{args.period.replace(" ", "")}_{stamp}.csv')
        table3_df.to_csv(t3_csv, index=False)
        print('Table 3 CSV:', t3_csv)

        sub_table2 = build_table2_subdoc(tpl, conc_df)
        table3_context = compute_table3_context(conc_df)
        
        fig4_path = os.path.join(args.outdir, f'figure4_concordance_{args.period.replace(" ", "")}_{stamp}.png')
        build_figure4(conc_df, fig4_path)

    context = {
        'figure2': InlineImage(tpl, fig_path, width=Mm(140)),
        'table1_indicators': sub_table1,
        'table2_concordance_facility': sub_table2,
        'figure4': '',  # Default empty if not generated
        **table3_context
    }
    
    if fig4_path and os.path.exists(fig4_path):
        context['figure4'] = InlineImage(tpl, fig4_path, width=Mm(140))
        
    tpl.render(context)

    out_docx = args.out or os.path.join(args.outdir, f'CDC_Report_{args.period.replace(" ", "")}_{stamp}.docx')
    tpl.save(out_docx)
    print('Figure 2 PNG:', fig_path)
    if fig4_path:
        print('Figure 4 PNG:', fig4_path)
    print('Table 1 CSV:', csv_path)
    print('Rendered report:', out_docx)

if __name__ == '__main__':
    main()
