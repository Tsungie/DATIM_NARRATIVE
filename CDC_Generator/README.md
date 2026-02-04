# 📊 Datim Report Generator (Datim Narrative)

This tool generates the standard Zim-TTECH CDC Quarterly Report (Word Document) automatically. 
It takes Excel data (Consistency & Concordance reports), creates the required charts (Figures 1, 2, and 4), calculates summary statistics, and fills them into a Word template.

## ✅ Prerequisites

You need **Python 3.8+** installed on your computer.

## 🚀 Installation

1.  **Clone or Download** this folder.
2.  **Create a Virtual Environment** (recommended):
    ```bash
    python -m venv venv
    source venv/bin/activate      # On Mac/Linux
    venv\Scripts\activate         # On Windows
    ```
3.  **Install Dependencies**:
    Create a file named `requirements.txt` with the list below, then run:
    ```bash
    pip install -r requirements.txt
    ```

### `requirements.txt`
```text
streamlit
pandas
matplotlib
seaborn
docxtpl
python-docx
openpyxl
xlsxwriter