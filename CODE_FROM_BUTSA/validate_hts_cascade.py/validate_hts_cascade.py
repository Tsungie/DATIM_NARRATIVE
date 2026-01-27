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


def validate_hts_cascade(input_file: str, output_file: str = None) -> pd.DataFrame:
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

    # Totals
    total_tested_sum = results_df["HTS_Total_Tested"].sum()
    total_positive_sum = results_df["HTS_Total_Positive"].sum()
    tx_new_sum = results_df["TX_NEW"].sum()

    print(f"\nOverall Totals:")
    print(f"  Total Tested: {total_tested_sum:,}")
    print(f"  Total Positive: {total_positive_sum:,}")
    print(f"  TX_NEW: {tx_new_sum:,}")
    print(
        f"  Positivity Rate: {round(total_positive_sum / total_tested_sum * 100, 1) if total_tested_sum > 0 else 0}%"
    )
    print(
        f"  Linkage Rate: {round(tx_new_sum / total_positive_sum * 100, 1) if total_positive_sum > 0 else 0}%"
    )

    if anomaly_count > 0:
        print(f"\n{'=' * 60}")
        print("FACILITIES WITH ANOMALIES (TX_NEW > Positives):")
        print(f"{'=' * 60}")
        anomalies = results_df[results_df["Validation_Status"] != "OK"]
        for _, row in anomalies.iterrows():
            print(f"  - {row['Facility']} ({row['District']})")
            print(
                f"    Positives: {row['HTS_Total_Positive']}, TX_NEW: {row['TX_NEW']}"
            )

    # Save to Excel if output file specified
    if output_file:
        print(f"\nSaving results to: {output_file}")

        # Create Excel with formatting
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            results_df.to_excel(writer, sheet_name="Facility Validation", index=False)

            # Add summary sheet
            summary_data = {
                "Metric": [
                    "Total Facilities",
                    "Facilities with Anomalies",
                    "Facilities OK",
                    "Total Tested",
                    "Total Positive",
                    "TX_NEW",
                    "Overall Positivity Rate",
                    "Overall Linkage Rate",
                    "Validation Date",
                    "Input File",
                ],
                "Value": [
                    total_facilities,
                    anomaly_count,
                    total_facilities - anomaly_count,
                    total_tested_sum,
                    total_positive_sum,
                    tx_new_sum,
                    f"{round(total_positive_sum / total_tested_sum * 100, 1)}%"
                    if total_tested_sum > 0
                    else "N/A",
                    f"{round(tx_new_sum / total_positive_sum * 100, 1)}%"
                    if total_positive_sum > 0
                    else "N/A",
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    os.path.basename(input_file),
                ],
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name="Summary", index=False)

            # Add anomalies-only sheet
            if anomaly_count > 0:
                anomalies_df = results_df[results_df["Validation_Status"] != "OK"]
                anomalies_df.to_excel(writer, sheet_name="Anomalies Only", index=False)

        print(f"Results saved successfully!")

    print(f"\n{'=' * 60}")

    return results_df


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("\nError: Please provide input DATIM file path")
        sys.exit(1)

    input_file = sys.argv[1]

    if not os.path.exists(input_file):
        print(f"Error: File not found: {input_file}")
        sys.exit(1)

    # Default output file
    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        # Generate output filename based on input
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        output_file = (
            f"validation_{base_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

    validate_hts_cascade(input_file, output_file)


if __name__ == "__main__":
    main()

# ⏺ The discrepancy is likely because the datim_report table has multiple periods or multiple generations. The query needs to filter by generated_at to get only the latest generation.

#   Let me help you check:

#   -- Check how many records exist for Avondale and what periods/generations
#   SELECT
#       period,
#       generated_at,
#       COUNT(*) as records,
#       SUM(value) as total_value
#   FROM datim_report
#   WHERE org_unit = 'LixFsP9YMT9'
#   GROUP BY period, generated_at
#   ORDER BY generated_at DESC;

#   Fixed query with generated_at filter:

#   -- HTS CASCADE for Avondale - LATEST GENERATION ONLY
#   SELECT
#       'Avondale - 100030 - Clinic' AS facility,

#       SUM(CASE WHEN datim_indicator IN (
#           'HTS_INDEX_COM_N_DSD_Age_Sex_Result',
#           'HTS_INDEX_FAC_N_TA_Age_Sex_Result',
#           'HTS_TST_N_TA_Inpat_Age_Sex_Result',
#           'HTS_TST_N_TA_OtPITC_Age_Sex_Result',
#           'HTS_TST_N_TA_PMTCT_PB_Age_Sex_Result',
#           'HTS_TST_N_TA_PMTCT_PP_LandD_Age_Sex_Result',
#           'HTS_TST_N_TA_SNS_Age_Sex_Result',
#           'HTS_TST_N_TA_STI_Age_Sex_Result',
#           'PMTCT_STAT_N_TA_Age_Sex_KnownNewResult',
#           'TB_STAT_N_TA_Age_Sex_KnownNewPosNeg'
#       ) THEN value ELSE 0 END) AS total_tested,

#       SUM(CASE WHEN datim_indicator IN (
#           'HTS_INDEX_COM_N_DSD_Age_Sex_Result',
#           'HTS_INDEX_FAC_N_TA_Age_Sex_Result',
#           'HTS_TST_N_TA_Inpat_Age_Sex_Result',
#           'HTS_TST_N_TA_OtPITC_Age_Sex_Result',
#           'HTS_TST_N_TA_PMTCT_PB_Age_Sex_Result',
#           'HTS_TST_N_TA_PMTCT_PP_LandD_Age_Sex_Result',
#           'HTS_TST_N_TA_SNS_Age_Sex_Result',
#           'HTS_TST_N_TA_STI_Age_Sex_Result',
#           'PMTCT_STAT_N_TA_Age_Sex_KnownNewResult',
#           'TB_STAT_N_TA_Age_Sex_KnownNewPosNeg'
#       ) AND element_result_group IN ('Positive', 'Newly Tested Positives', 'New Positives')
#       THEN value ELSE 0 END) AS total_positive,

#       SUM(CASE WHEN datim_indicator = 'TX_NEW_N_TA_Age_Sex_HIVStatus'
#           THEN value ELSE 0 END) AS tx_new

#   FROM datim_report
#   WHERE period = 'Cop25 Q1'
#   AND org_unit = 'LixFsP9YMT9'
#   AND generated_at = (
#       SELECT MAX(generated_at)
#       FROM datim_report
#       WHERE period = 'Cop25 Q1'
#   );
