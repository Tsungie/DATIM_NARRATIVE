from get_datim_data import fetch_data as datim
from get_dhis2_data import fetch_data as dhis2
import pandas as pd
import numpy as np


def main():
    df_datim = datim("input_files/DATIM_Report.xlsx")
    df_dhis2 = dhis2()

    print(df_datim.columns)
    print(df_dhis2.columns)


    df_merge = pd.merge(df_datim,df_dhis2, on = 'Facility ID', how = 'left', suffixes=[' datim', ' dhis2'])

    # Define pairs of columns to compare for concordance
    indicator_pairs = {
        "HTS_TST": ("HTS_TST datim", "HTS_TST dhis2"),
        "HTS_POS": ("HTS_POS datim", "HTS_POS dhis2"),
        "TX_NEW": ("TX_NEW datim", "TX_NEW dhis2"),
    }

    # Calculate concordance for each pair
    for indicator, (datim_col, dhis2_col) in indicator_pairs.items():
        if datim_col in df_merge.columns and dhis2_col in df_merge.columns:
            # Fill NaN with 0 for calculation
            df_merge[datim_col] = df_merge[datim_col].fillna(0)
            df_merge[dhis2_col] = df_merge[dhis2_col].fillna(0)

            numerator = np.minimum(df_merge[datim_col], df_merge[dhis2_col])
            denominator = np.maximum(df_merge[datim_col], df_merge[dhis2_col])
            
            concordance_col_name = f"{indicator}_Concordance"
            
            # If denominator is 0, both are 0, so they are 100% concordant.
            df_merge[concordance_col_name] = np.where(
                denominator == 0, 
                100.0, 
                (numerator / denominator) * 100
            )
            df_merge[concordance_col_name] = df_merge[concordance_col_name].round(0)

    # Rearrange columns to group indicators with their concordance

    df_merge.rename(columns={"Province datim": "Province", "District datim": "District", "Facility datim":"Facility "})
    id_cols = [col for col in ['Province', 'District', 'Facility'] if col in df_merge.columns]
    indicator_cols_ordered = []
    for indicator in ["HTS_TST", "HTS_POS", "TX_NEW"]:
        group = [f"{indicator} datim", f"{indicator} dhis2", f"{indicator}_Concordance"]
        indicator_cols_ordered.extend([col for col in group if col in df_merge.columns])

    final_column_order = id_cols + indicator_cols_ordered
    df_merge = df_merge[final_column_order]
    
    df_merge.to_csv("df_merge.csv", index=False)

if __name__ == "__main__":
    main()