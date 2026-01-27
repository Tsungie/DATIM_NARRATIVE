#!/usr/bin/env python3
"""
HTS Cascade Validation Script
Validates that TX_NEW <= HTS Positives per facility

Usage:
    python scripts/validate_hts_cascade.py <datim_file.xlsx> [output_file.xlsx]

Example:
    python scripts/validate_hts_cascade.py "static/DATIM_Report_COP2025_Q1 7.xlsx" "validation_results.xlsx"
"""

import pandas as pd
import sys
import os
from datetime import datetime

# HTS Cascade Indicator Definitions (matching our custom_indicators)
HTS_TESTED_INDICATORS = [
    "HTS_INDEX_COM_N_DSD_Age_Sex_Result",
    "HTS_INDEX_FAC_N_TA_Age_Sex_Result",
    "HTS_TST_N_TA_Inpat_Age_Sex_Result",
    "HTS_TST_N_TA_OtPITC_Age_Sex_Result",
    "HTS_TST_N_TA_PMTCT_PB_Age_Sex_Result",
    "HTS_TST_N_TA_PMTCT_PP_LandD_Age_Sex_Result",
    "HTS_TST_N_TA_SNS_Age_Sex_Result",
    "HTS_TST_N_TA_STI_Age_Sex_Result",
    "PMTCT_STAT_N_TA_Age_Sex_KnownNewResult",
    "TB_STAT_N_TA_Age_Sex_KnownNewPosNeg",
]

# Result groups that indicate POSITIVE results
POSITIVE_RESULT_GROUPS = [
    "Positive",
    "Positives",
    "Newly Tested Positives",
    "New Positives",
    "New Positive",
]

# TX_NEW Indicator
TX_NEW_INDICATOR = "TX_NEW_N_TA_Age_Sex_HIVStatus"


def fetch_data(input_file) -> pd.DataFrame:
    """
    Validate HTS Cascade data from DATIM file.

    Args:
        input_file: Path to DATIM Excel file
        output_file: Optional path to output Excel file

    Returns:
        DataFrame with validation results
    """
    print(f"\n{'=' * 60}")
    print("HTS CASCADE VALIDATION SCRIPT")
    print(f"{'=' * 60}")
    print(f"Input file: {input_file}")

    # Read DATIM file
    print("\nReading DATIM file...")
    df = pd.read_excel(input_file)
    print(f"Total rows: {len(df):,}")
    print(f"Columns: {df.columns.tolist()}")

    # Standardize column names (handle variations)
    col_mapping = {
        "SiteProvince": "Province",
        "SiteDistrict": "District",
        "SiteName": "Facility",
        "DATIM_Indicator": "DATIM_Indicator",
        "ElementResultGroup": "ResultGroup",
        "Value": "Value",
        "Org_Unit": "Org_Unit",
    }

    # Rename columns that exist
    for old_col, new_col in col_mapping.items():
        if old_col in df.columns and old_col != new_col:
            df = df.rename(columns={old_col: new_col})

    # Ensure Value is numeric
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0)

    # Get unique facilities
    facilities = df[["Province", "District", "Facility", "Org_Unit"]].drop_duplicates()
    print(f"Unique facilities: {len(facilities)}")

    results = []

    for _, facility in facilities.iterrows():
        province = facility["Province"]
        district = facility["District"]
        facility_name = facility["Facility"]
        org_unit = facility["Org_Unit"]

        # Filter data for this facility
        facility_data = df[df["Org_Unit"] == org_unit]

        # Calculate HTS Total Tested (all results from HTS indicators)
        hts_tested_data = facility_data[
            facility_data["DATIM_Indicator"].isin(HTS_TESTED_INDICATORS)
        ]
        total_tested = hts_tested_data["Value"].sum()

        # Calculate HTS Total Positive (only positive results)
        hts_positive_data = facility_data[
            (facility_data["DATIM_Indicator"].isin(HTS_TESTED_INDICATORS))
            & (facility_data["ResultGroup"].isin(POSITIVE_RESULT_GROUPS))
        ]
        total_positive = hts_positive_data["Value"].sum()

        # Calculate TX_NEW
        tx_new_data = facility_data[
            facility_data["DATIM_Indicator"] == TX_NEW_INDICATOR
        ]
        tx_new = tx_new_data["Value"].sum()

        # Determine anomaly
        has_anomaly = tx_new > total_positive
        if has_anomaly:
            gap = tx_new - total_positive
            comment = f"ANOMALY: TX_NEW ({int(tx_new)}) > Positives ({int(total_positive)}) by {int(gap)}"
        else:
            comment = "OK"

        results.append(
            {
                "Province": province,
                "District": district,
                "Facility": facility_name,
                "Org_Unit": org_unit,
                "HTS_Total_Tested": int(total_tested),
                "HTS_Total_Positive": int(total_positive),
                "TX_NEW": int(tx_new),
                "Positivity_Rate": round((total_positive / total_tested * 100), 1)
                if total_tested > 0
                else 0,
                "Linkage_Rate": round((tx_new / total_positive * 100), 1)
                if total_positive > 0
                else 0,
                "Validation_Status": comment,
            }
        )

    # Create results DataFrame
    results_df = pd.DataFrame(results)

    # Sort by Province, District, Facility
    results_df = results_df.sort_values(["Province", "District", "Facility"])

    # Print summary
    print(f"\n{'=' * 60}")
    print("VALIDATION SUMMARY")
    print(f"{'=' * 60}")

    total_facilities = len(results_df)
    anomaly_count = len(results_df[results_df["Validation_Status"] != "OK"])

    print(f"Total Facilities: {total_facilities}")
    print(f"Facilities with Anomalies: {anomaly_count}")
    print(f"Facilities OK: {total_facilities - anomaly_count}")

    results_df.rename(columns={"Org_Unit": "Facility ID",
                               "HTS_Total_Tested":"HTS_TST",
                               "HTS_Total_Positive":"HTS_POS"
                               }, inplace=True)

    results_df = results_df[[
       "Province", "District",	"Facility"	,"Facility ID"	,"HTS_TST"	,"HTS_POS",	"TX_NEW"
    ]]
    print("succeful")

    results_df.to_csv("datim_1.csv", index=False)

    return results_df


