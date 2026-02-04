import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
from datetime import datetime
import io

# --- 1. CONFIGURATION & MAPPINGS ---
st.set_page_config(page_title="Datim  Report Generator", layout="wide")

TARGETS = {
    'TX_CURR': 823, 'TX_TB': 823, 'TX_ML': 823, 'HTS_TST and HTS_POS': 823,
    'PMTCT_STAT': 761, 'PMTCT_ART': 659, 'TB_PREV': 657, 'TX_PVLS': 823,
    'HTS_INDEX': 823, 'TX_NEW': 823, 'TB_ART': 657, 'PREP_NEW': 628,
    'PMTCT_FO': 661, 'HTS_SELF': 823, 'TB_STAT': 657, 'PMTCT_EID': 655,
    'CXCA_SCRN': 628, 'CXCA_TX': 473, 'PREP_CT': 628, 'TX_RTT': 823,
    'PMTCT_HEI': 661,
}

MAP_RULES = {
    'TX_CURR': ['TX_CURR TA'], 'TX_TB': ['TX_TB(DENOM) TA', 'TX_TB TA'],
    'TX_ML': ['TX_ML TA'], 'HTS_TST': ['HTS_TST TA'], 'HTS_POS': ['HTS_POS TA'],
    'PMTCT_STAT': ['PMTCT_STAT(DENOM) TA', 'PMTCT_STAT TA'], 'PMTCT_ART': ['PMTCT_ART TA'],
    'TB_PREV': ['TB_PREV(DENOM) TA', 'TB_PREV TA'], 'TX_PVLS': ['TX_PVLS(DENOM) TA', 'TX_PVLS TA'],
    'HTS_INDEX': ['HTS_INDEX TA'], 'TX_NEW': ['TX_NEW TA'], 'TB_ART': ['TB_ART TA'],
    'PREP_NEW': ['PrEP_NEW TA', 'PREP_NEW TA'], 'PMTCT_FO': ['PMTCT_FO TA'],
    'HTS_SELF': ['HTS_SELF TA'], 'TB_STAT': ['TB_STAT(DENOM) TA', 'TB_STAT TA'],
    'PMTCT_EID': ['PMTCT_EID TA'], 'CXCA_SCRN': ['CXCA_SCRN TA'], 'CXCA_TX': ['CXCA_TX TA'],
    'PREP_CT': ['PrEP_CT TA', 'PREP_CT TA'], 'TX_RTT': ['TX_RTT TA'], 'PMTCT_HEI': ['PMTCT_HEI TA'],
}

INDICATOR_ORDER = [
    'TX_CURR', 'TX_TB', 'TX_ML', 'HTS_TST and HTS_POS', 'PMTCT_STAT', 'PMTCT_ART',
    'TB_PREV', 'TX_PVLS', 'HTS_INDEX', 'TX_NEW', 'TB_ART', 'PREP_NEW', 'PMTCT_FO',
    'HTS_SELF', 'TB_STAT', 'PMTCT_EID', 'CXCA_SCRN', 'CXCA_TX', 'PREP_CT', 'TX_RTT', 'PMTCT_HEI'
]

CONCORDANCE_MAP = {
    "HTS_TST":  ("MRF HTS_TX_NEW (Total Tests)",        "DATIM HTS_TX_NEW (Total Tests)"),
    "HTS_POS":  ("MRF HTS_TX_NEW (Total Positives)",    "DATIM HTS_TX_NEW (Total Positives)"),
    "TX_NEW":   ("MRF HTS_TX_NEW (Total Initiations)",  "DATIM HTS_TX_NEW (Total Initiations)"),
    "TX_CURR":  ("MRF TX_CURR",                         "DATIM TX_CURR"),
}

# --- 2. HELPER FUNCTIONS ---

def load_consistency(file):
    df = pd.read_excel(file, engine='openpyxl', sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    ind_col = next((c for c in df.columns if 'indicator' in str(c).lower()), df.columns[0])
    num_col = next((c for c in df.columns if ('number' in str(c).lower() or 'no.' in str(c).lower()) and 'report' in str(c).lower()), df.columns[1])
    return pd.DataFrame({
        'Indicator': df[ind_col].astype(str).str.strip(),
        'ReportingUnits': pd.to_numeric(df[num_col], errors='coerce').fillna(0).astype(int)
    })

def build_21_indicator_table(cons_df):
    rows = []
    values = {}
    for key, pats in MAP_RULES.items():
        # Pick reporting
        found = 0
        for pat in pats:
            row = cons_df[cons_df['Indicator'] == pat]
            if not row.empty:
                found = int(row.iloc[0]['ReportingUnits'])
                break
        values[key] = found
        
    hts_combined = max(values.get('HTS_TST', 0), values.get('HTS_POS', 0))
    for ind in INDICATOR_ORDER:
        reporting = hts_combined if ind == 'HTS_TST and HTS_POS' else values.get(ind, 0)
        tgt = TARGETS[ind]
        comp = round(100 * reporting / tgt, 1) if tgt else 0.0
        rows.append({'Indicator': ind, 'Number_of_facilities_reporting': reporting, 'Target': tgt, 'Reporting_Completeness_%': comp})
    return pd.DataFrame(rows)

def load_concordance(file):
    raw = pd.read_excel(file, engine='openpyxl')
    raw.columns = [str(c).strip() for c in raw.columns]
    rows = []
    for ind, (mrf_col, ehr_col) in CONCORDANCE_MAP.items():
        if mrf_col not in raw.columns or ehr_col not in raw.columns: continue
        tmp = raw[["Facility", "Province", "District", mrf_col, ehr_col]].copy()
        tmp["Indicator"] = ind
        tmp["MRF"] = pd.to_numeric(tmp[mrf_col], errors="coerce").fillna(0)
        tmp["EHR"] = pd.to_numeric(tmp[ehr_col], errors="coerce").fillna(0)
        tmp["Concordance_%"] = tmp.apply(lambda r: (100.0 if r["MRF"]==0 and r["EHR"]==0 else (0.0 if r["MRF"]==0 else max(0.0, min(100.0, (1 - abs(r["EHR"] - r["MRF"]) / r["MRF"]) * 100)))), axis=1)
        rows.append(tmp[["Facility", "Province", "District", "Indicator", "MRF", "EHR", "Concordance_%"]])
    return pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()

def build_figure1_cascade(total, manual, pipeline, mobile, analyzed):
    # Calculations
    total_collected = manual + pipeline + mobile
    not_collected = total - total_collected
    
    # Percentages
    coll_pct = (total_collected / total * 100) if total > 0 else 0
    analyzed_pct = (analyzed / total_collected * 100) if total_collected > 0 else 0
    total_failed = total_collected - analyzed
    failed_pct = (total_failed / total_collected * 100) if total_collected > 0 else 0

    # Plotting
    fig, ax = plt.subplots(figsize=(16, 14))
    ax.set_xlim(0, 100); ax.set_ylim(0, 100); ax.axis('off')
    
    # Helper to format
    def fmt(val, parent): return f"{int(val)} ({(val/parent*100 if parent>0 else 0):.0f}%)"
    
    # Helper box
    def add_box(xy, text, bg, fs=10, style="round,pad=1"):
        ax.text(xy[0], xy[1], text, ha='center', va='center', fontsize=fs, fontweight='bold',
                bbox=dict(boxstyle=style, fc=bg, ec="black", lw=1))

    # Helper line
    def connect(xy1, xy2): ax.plot([xy1[0], xy2[0]], [xy1[1], xy2[1]], color='#555555', lw=1.5, zorder=0)

    # Coordinates
    xy_total=(50, 96); xy_coll=(30, 84); xy_not=(70, 84)
    xy_man=(16, 62); xy_mob=(50, 62); xy_dpl=(84, 62)
    xy_man_S=(8, 32); xy_man_F=(24, 32); xy_mob_S=(42, 32); xy_mob_F=(58, 32); xy_dpl_S=(76, 32); xy_dpl_F=(92, 32)

    # Draw lines
    connect(xy_total, xy_coll); connect(xy_total, xy_not)
    connect(xy_coll, xy_man); connect(xy_coll, xy_mob); connect(xy_coll, xy_dpl)
    connect(xy_man, xy_man_S); connect(xy_man, xy_man_F)
    connect(xy_mob, xy_mob_S); connect(xy_mob, xy_mob_F)
    connect(xy_dpl, xy_dpl_S); connect(xy_dpl, xy_dpl_F)

    # Draw Nodes
    add_box(xy_total, f"Total Full Impilo EHR Sites\n{total}", '#A0CBE8', 12)
    add_box(xy_coll, f"Backups Collected\n{fmt(total_collected, total)}", '#A0CBE8')
    add_box(xy_not, f"Backups NOT Collected\n{fmt(not_collected, total)}", '#FADBD8')
    add_box(xy_man, f"Physical Backups\nCollected\n{fmt(manual, total_collected)}", '#FEF5E7')
    add_box(xy_mob, f"Mobile App\nData\n{fmt(mobile, total_collected)}", '#FEF5E7')
    add_box(xy_dpl, f"DPL / Pipeline\nData\n{fmt(pipeline, total_collected)}", '#FEF5E7')

    # Success/Fail estimation logic (proportional)
    def split(src):
        if total_collected == 0: return 0,0
        s = int(src * (analyzed/total_collected)); return s, src-s
    
    ms, mf = split(manual); bs, bf = split(mobile); ps, pf = split(pipeline)
    # Fix rounding
    ms += (analyzed - (ms+bs+ps))

    lbl_s="Extracted and available for\nanalysis and reporting"
    lbl_f="Backups that failed\nto process"
    
    for xy, val, col, lbl in [(xy_man_S, ms, '#EBF7EB', lbl_s), (xy_man_F, mf, '#FADBD8', lbl_f),
                              (xy_mob_S, bs, '#EBF7EB', lbl_s), (xy_mob_F, bf, '#FADBD8', lbl_f),
                              (xy_dpl_S, ps, '#EBF7EB', lbl_s), (xy_dpl_F, pf, '#FADBD8', lbl_f)]:
        add_box(xy, f"{lbl}\n{fmt(val, (manual if xy in [xy_man_S, xy_man_F] else (mobile if xy in [xy_mob_S, xy_mob_F] else pipeline)))}", col, 8, "square,pad=0.8")

    # Summary
    summary = (f"SUMMARY STATISTICS\n────────────────────────────────────────\n"
               f"Collection Rate: {coll_pct:.1f}% ({total_collected}/{total})\n"
               f"Total Successfully Extracted: {analyzed_pct:.1f}% ({analyzed}/{total_collected})\n"
               f"Total Failed to Process: {failed_pct:.1f}% ({total_failed}/{total_collected})")
    ax.text(50, 8, summary, ha='center', va='center', fontsize=12, fontfamily='monospace',
            bbox=dict(boxstyle="round,pad=1.2", fc='#FAFAFA', ec='black', lw=1.5))
    
    plt.title('Figure 1: Summary of IMPILO E-HR Data Flow Cascade', fontsize=16, fontweight='bold', y=0.99)
    return fig

def build_figure2(table_df, period):
    sns.set(style='whitegrid')
    fig, ax = plt.subplots(figsize=(12,8))
    plot_df = table_df.sort_values('Reporting_Completeness_%', ascending=False)
    sns.barplot(x='Indicator', y='Reporting_Completeness_%', data=plot_df, color='#4C78A8', ax=ax)
    ax.set_ylabel('Reporting Completeness (%)'); ax.set_xlabel('Indicator'); ax.set_ylim(0,100)
    ax.set_title(f'Figure 2: Proportion of sites reporting individual indicators in IMPILO E-HR {period}')
    ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')
    for p in ax.patches:
        ax.annotate(f"{p.get_height()}%", (p.get_x()+p.get_width()/2, p.get_height()), ha='center', va='bottom', fontsize=9)
    plt.tight_layout()
    return fig

def build_figure4(df):
    target_provinces = ['Harare', 'Bulawayo']
    mask = df['Province'].astype(str).str.strip().apply(lambda x: any(p.lower() in x.lower() for p in target_provinces))
    mask_ind = df['Indicator'].astype(str).str.upper() == 'TX_CURR'
    subset = df[mask & mask_ind].copy()
    if subset.empty: return None

    subset['Province'] = subset['Province'].apply(lambda p: 'Harare' if 'harare' in str(p).lower() else ('Bulawayo' if 'bulawayo' in str(p).lower() else p))
    grouped = subset.groupby('Province')[['MRF', 'EHR']].sum().reset_index()
    total_row = pd.DataFrame({'Province': ['Total'], 'MRF': [grouped['MRF'].sum()], 'EHR': [grouped['EHR'].sum()]})
    ordered = pd.concat([total_row, grouped], ignore_index=True)
    ordered['Concordance'] = ordered.apply(lambda r: (r['EHR']/r['MRF']*100) if r['MRF']>0 else 0, axis=1)
    
    plot_df = ordered.melt(id_vars='Province', value_vars=['MRF', 'EHR'], var_name='Source', value_name='Count')
    x_order = ['Total'] + [p for p in target_provinces if p in grouped['Province'].unique()]

    sns.set(style='white'); fig, ax1 = plt.subplots(figsize=(11, 6))
    sns.barplot(x='Province', y='Count', hue='Source', data=plot_df, order=x_order, palette={'MRF':'#104E8B','EHR':'#00FF00'}, edgecolor="black", ax=ax1)
    if ax1.get_legend(): ax1.get_legend().remove()
    
    ax1.set_ylabel('Number of clients on ART'); ax1.set_title('TX_CURR Concordance: MRF vs EHR', fontsize=14, pad=20)
    for c in ax1.containers: ax1.bar_label(c, fmt='{:,.0f}', padding=3, fontsize=10)
    
    ax2 = ax1.twinx(); ax2.set_ylim(0, 115); ax2.set_yticks([])
    y_vals = [ordered.loc[ordered['Province']==r, 'Concordance'].values[0] for r in x_order]
    x_coords = range(len(x_order))
    ax2.plot(x_coords, y_vals, color='#D62728', marker='D', markersize=8, linestyle='None')
    for x, y in zip(x_coords, y_vals): ax2.text(x, y+4, f"{y:.1f}%", color='#D62728', ha='center', fontweight='bold')
    
    ax1.spines['top'].set_visible(False); ax1.spines['right'].set_visible(False); ax2.spines['top'].set_visible(False); ax2.spines['right'].set_visible(False)
    
    legend_elements = [mpatches.Patch(facecolor='#104E8B', edgecolor='black', label='MRF'), mpatches.Patch(facecolor='#00FF00', edgecolor='black', label='EHR'), Line2D([0],[0], color='#D62728', marker='D', linestyle='None', markersize=8, label='Concordance')]
    ax1.legend(handles=legend_elements, title='Key', bbox_to_anchor=(1.02, 1), loc='upper left')
    plt.tight_layout()
    return fig

# --- 3. DOCXTPL FILLER FUNCTIONS ---
def shade_cell(cell, color):
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), color); cell._tc.get_or_add_tcPr().append(shd)

def build_tables(tpl, df_tabs, type_tab):
    sub = tpl.new_subdoc()
    cols = 4 if type_tab==1 else 7
    tbl = sub.add_table(rows=1, cols=cols)
    try: tbl.style = 'Table Grid'
    except: pass
    
    headers = ['Indicator','Number of facilities reporting','Target','Reporting Completeness'] if type_tab==1 else ["Facility","Province","District","Indicator","MRF","EHR","Concordance_%"]
    for i, h in enumerate(headers): tbl.rows[0].cells[i].text=h; [run.bold for p in tbl.rows[0].cells[i].paragraphs for run in p.runs]
    
    for _, r in df_tabs.iterrows():
        row = tbl.add_row().cells
        vals = [r['Indicator'], r['Number_of_facilities_reporting'], r['Target'], f"{r['Reporting_Completeness_%']}%"] if type_tab==1 else [r.get(h, "") for h in headers]
        for i, v in enumerate(vals):
            if type_tab==2 and headers[i]=="Concordance_%":
                try: 
                    fv = float(v)
                    row[i].text = f"{fv:.1f}%"
                    color = 'C6EFCE' if fv>=95 else ('FFEB9C' if fv>=90 else 'F8CBAD')
                    shade_cell(row[i], color)
                except: row[i].text=""
            else:
                row[i].text = str(v)
    return sub

# --- 4. MAIN STREAMLIT APP ---
st.title("📊 CDC Report Generator (Datim Narrative)")

col1, col2 = st.columns([1, 2])

with col1:
    st.header("1. Upload Files")
    f_consistency = st.file_uploader("Consistency Report (Excel)", type=['xlsx'])
    f_concordance = st.file_uploader("Concordance Analysis (Excel)", type=['xlsx'])
    f_template = st.file_uploader("Word Template (docx)", type=['docx'])
    
    st.header("2. Settings")
    period = st.text_input("Period Label", "COP24 Q4")
    
    with st.expander("3. Cascade Data Numbers (Figure 1)", expanded=True):
        total_sites = st.number_input("Total Active Sites", value=823)
        manual = st.number_input("Collected Manually", value=322)
        pipeline = st.number_input("Pipeline / DPL", value=81)
        mobile = st.number_input("Mobile Backups", value=170)
        analyzed = st.number_input("Successfully Analyzed", value=449)
        
    btn_generate = st.button("Generate Report", type="primary")

with col2:
    if btn_generate:
        if not f_consistency or not f_template:
            st.error("Please upload at least the Consistency Report and the Word Template.")
        else:
            with st.spinner("Crunching numbers and painting charts..."):
                # 1. Processing Consistency
                cons_df = load_consistency(f_consistency)
                table1_df = build_21_indicator_table(cons_df)
                fig2 = build_figure2(table1_df, period)
                
                # 2. Processing Cascade (Fig 1)
                fig1 = build_figure1_cascade(total_sites, manual, pipeline, mobile, analyzed)
                
                # 3. Processing Concordance (if uploaded)
                sub_table2 = ""
                table3_ctx = {}
                fig4 = None
                
                # Template Prep
                tpl = DocxTemplate(f_template)
                
                # Save figures temporarily to disk for docxtpl (it requires paths)
                if not os.path.exists("temp_figs"): os.makedirs("temp_figs")
                fig1.savefig("temp_figs/fig1.png", dpi=200, bbox_inches='tight')
                fig2.savefig("temp_figs/fig2.png", dpi=200, bbox_inches='tight')
                
                if f_concordance:
                    conc_df = load_concordance(f_concordance)
                    sub_table2 = build_tables(tpl, conc_df, 2)
                    fig4 = build_figure4(conc_df)
                    if fig4: fig4.savefig("temp_figs/fig4.png", dpi=200, bbox_inches='tight')
                    
                    # Table 3 Context
                    for ind in ['HTS_TST', 'HTS_POS', 'TX_NEW', 'TX_CURR']:
                        sub = conc_df[conc_df['Indicator'].astype(str).str.upper() == ind]
                        m = sub['MRF'].sum(); e = sub['EHR'].sum()
                        c = (e/m*100) if m>0 else 0
                        table3_ctx[f"overall_{ind.lower()}_mrf"] = f"{int(m):,}"
                        table3_ctx[f"overall_{ind.lower()}_ehr"] = f"{int(e):,}"
                        table3_ctx[f"overall_{ind.lower()}_conc"] = f"{c:.1f}"

                # Render
                context = {
                    'total_active_facilities': total_sites,
                    'collected_manually': manual, 'pushed_via_pipeline': pipeline, 'mobile_backups': mobile,
                    'ehr_facilities_analyzed': analyzed, 
                    'ehr_facilities_analyzed_pct': round(analyzed/total_sites*100, 1),
                    'figure1': InlineImage(tpl, "temp_figs/fig1.png", width=Mm(140)),
                    'figure2': InlineImage(tpl, "temp_figs/fig2.png", width=Mm(140)),
                    'table1_indicators': build_tables(tpl, table1_df, 1),
                    'table2_concordance_facility': sub_table2,
                    'figure4': InlineImage(tpl, "temp_figs/fig4.png", width=Mm(140)) if fig4 else "",
                    **table3_ctx
                }
                
                tpl.render(context)
                
                # Save to buffer
                bio = io.BytesIO()
                tpl.save(bio)
                
                st.success("Report Generated Successfully!")
                
                # Preview Images
                st.subheader("Preview Charts")
                c1, c2 = st.columns(2)
                c1.image("temp_figs/fig1.png", caption="Figure 1: Cascade")
                c2.image("temp_figs/fig2.png", caption="Figure 2: Completeness")
                if fig4: st.image("temp_figs/fig4.png", caption="Figure 4: Concordance")
                
                # Download Button
                st.download_button(
                    label="📥 Download Final Word Report",
                    data=bio.getvalue(),
                    file_name=f"CDC_Report_{period.replace(' ','_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )