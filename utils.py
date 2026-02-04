import pandas as pd
import numpy as np
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# -----------------------------
# Indicator mapping for Concordance Analysis
# -----------------------------
CONCORDANCE_MAP = {
    "HTS_TST":  ("MRF HTS_TX_NEW (Total Tests)",        "DATIM HTS_TX_NEW (Total Tests)"),
    "HTS_POS":  ("MRF HTS_TX_NEW (Total Positives)",    "DATIM HTS_TX_NEW (Total Positives)"),
    "TX_NEW":   ("MRF HTS_TX_NEW (Total Initiations)",  "DATIM HTS_TX_NEW (Total Initiations)"),
    "TX_CURR":  ("MRF TX_CURR",                         "DATIM TX_CURR"),
}

# -----------------------------
# Provinces for Harare vs Bulawayo analysis
# -----------------------------
HB_PROVINCES = {"Harare Metropolitan", "Bulawayo Metropolitan"}

# -----------------------------
# Default Indicator targets for reporting completeness (Figure 2)
# -----------------------------
DEFAULT_TARGETS = {
    "TX_CURR": 823,
    "TX_TB": 823,
    "TX_ML": 823,
    "HTS_TST and HTS_POS": 823,
    "PMTCT_STAT": 761,
    "PMTCT_ART": 659,
    "TB_PREV": 657,
    "TX_PVLS": 823,
    "HTS_INDEX": 823,
    "TX_NEW": 823,
    "TB_ART": 657,
    "PREP_NEW": 628,
    "PMTCT_FO": 661,
    "HTS_SELF": 823,
    "TB_STAT": 657,
    "PMTCT_EID": 655,
    "CXCA_SCRN": 628,
    "CXCA_TX": 473,
    "PREP_CT": 628,
    "TX_RTT": 823,
    "PMTCT_HEI": 661,
}

# -----------------------------
# Concordance functions (CDC-consistent)
# -----------------------------
def concordance_value(mrf, ehr):
    """
    Facility-level concordance (%), CDC-style:
      - If MRF=0 and EHR=0 => 100
      - If MRF=0 and EHR>0 => 0
      - Else: (1 - |EHR - MRF| / MRF) * 100, clipped to [0, 100]
    """
    if pd.isna(mrf) or pd.isna(ehr):
        return np.nan
    if mrf == 0 and ehr == 0:
        return 100.0
    if mrf == 0 and ehr > 0:
        return 0.0
    pct = (1.0 - abs(ehr - mrf) / mrf) * 100.0
    return float(max(0.0, min(100.0, pct)))

def concordance_overall(mrf_series, ehr_series):
    """
    Overall (aggregate) concordance (%), CDC-style:
      (1 - sum(|EHR_i - MRF_i|) / sum(MRF_i)) * 100  if sum(MRF) > 0
      else 100 when all zeros.
    """
    mrf_sum = mrf_series.sum()
    if mrf_sum == 0:
        return 100.0
    diff_sum = (ehr_series - mrf_series).abs().sum()
    pct = (1.0 - diff_sum / mrf_sum) * 100.0
    return float(max(0.0, min(100.0, pct)))

# -----------------------------
# Styling helpers for color-coding
# -----------------------------
def shade_cell(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def color_for_conc(pct):
    if pd.isna(pct):
        return None
    p = float(pct)
    if p >= 95.0:
        return 'C6EFCE'  # green
    if 90.0 <= p < 95.0:
        return 'FFEB9C'  # amber
    return 'F8CBAD'      # red

# -----------------------------
# Target loading for reporting completeness
# -----------------------------
def load_targets(targets_csv: str | None) -> dict:
    if targets_csv is None or not os.path.exists(targets_csv):
        return DEFAULT_TARGETS
    df = pd.read_csv(targets_csv)
    if not {"Indicator", "Target"}.issubset(df.columns):
        raise ValueError("Targets CSV must have columns: 'Indicator', 'Target'")
    return dict(zip(df["Indicator"], df["Target"]))