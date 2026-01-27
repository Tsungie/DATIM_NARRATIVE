from get_datim_data import fetch_data as datim
from get_dhis2_data import fetch_data as dhis2
import pandas as pd

df_datim = datim
df_dhis2 = dhis2

print(df_datim.columns)
print(df_dhis2.columns)


df_merge = pd.merge(df_datim,df_dhis2, on = 'Facility ID', how = 'left', suffixes=[' datim', ' dhis2'])
df_merge.to_csv("df_merge.csv")