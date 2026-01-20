# CDC Report Generator

This tool automates the creation of the CDC Narrative Report (Word document) by processing data from Consistency and Concordance Excel reports. It generates the required figures, tables, and CSV data exports (Figure 2, Table 1, Table 2, Table 3, Figure 4) and injects them directly into a Word template.

## Prerequisites

Ensure you have Python 3 installed. Install the required dependencies using pip:

```bash
pip install pandas matplotlib seaborn docxtpl python-docx openpyxl
```

## Usage

To generate the report, run the following command in your terminal:

```bash
python generate_cdc_report_docxtpl.py \
  --consistency "Consistency_Report_COP2024_Q4.xlsx" \
  --concordance "Concordance_Analysis_Complete.xlsx" \
  --template "cdc_template.docx" \
  --period "COP24 Q4" \
  --outdir "output"
```

### Arguments
*   `--consistency`: Path to the Consistency Report Excel file.
*   `--concordance`: Path to the Concordance Analysis Excel file.
*   `--template`: Path to the Word template file (default: `cdc_template.docx`).
*   `--period`: The reporting period label (e.g., "COP24 Q4") used in chart titles.
*   `--outdir`: Directory where the final DOCX, images, and CSVs will be saved.

## Outputs

The script generates the following files in the `output` directory:
1.  **CDC Report Word Doc**: `CDC_Report_<Period>_<Date>.docx` (The final narrative).
2.  **Figure 2**: `figure2_reporting_completeness_...png` (Bar chart).
3.  **Figure 4**: `figure4_concordance_...png` (Combo chart for Harare/Bulawayo).
4.  **Table 1 Data**: `figure2_table1_...csv` (Reporting completeness data).
5.  **Table 2 Data**: `table2_concordance_...csv` (Facility-level concordance).
6.  **Table 3 Data**: `table3_overall_concordance_...csv` (Overall concordance by indicator).

## How to Ensure Tables and Figures Generate Correctly

For the script to successfully read your data and produce the output, your input Excel files must meet the following criteria:

### 1. Consistency Report (`--consistency`)
Used to generate **Figure 2** and **Table 1**.
*   **Format**: Excel file (`.xlsx`).
*   **Required Columns**:
    *   `Indicator`: A column listing indicator names (e.g., "TX_CURR", "TX_NEW").
    *   `Number of Reporting Units`: A column containing the count of facilities reporting.
*   **Note**: The script looks for specific indicator names defined in the code (e.g., `TX_CURR TA`, `HTS_TST TA`). Ensure your indicator names match standard DATIM naming conventions.

### 2. Concordance Analysis (`--concordance`)
Used to generate **Table 2**, **Table 3**, and **Figure 4**.
*   **Format**: Excel file (`.xlsx`).
*   **Required Columns**:
    *   `Facility`, `Province`, `District`: For identifying sites.
    *   **Data Pairs**: The file must contain paired columns for MRF and DATIM/EHR data. For example:
        *   `MRF TX_CURR` and `DATIM TX_CURR`
        *   `MRF HTS_TX_NEW (Total Tests)` and `DATIM HTS_TX_NEW (Total Tests)`

### 3. Word Template (`--template`)
*   The template must be a `.docx` file.
*   It must contain **Jinja2 tags** where the content should be inserted. Ensure these tags exist in your document:
    *   `{{ figure2 }}`
    *   `{{ table1_indicators }}`
    *   `{{ table2_concordance_facility }}`
    *   `{{ figure4 }}`
    *   `{{ overall_tx_curr_mrf }}`, `{{ overall_tx_curr_ehr }}`, `{{ overall_tx_curr_conc }}`, etc. (for Table 3).

If any of these columns or tags are missing, the script may fail or produce empty sections in the report.