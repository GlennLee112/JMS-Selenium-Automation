import pandas as pd
import os
import glob

path = os.getcwd()
DP_folder = os.path.join(path,'DP List')
DP_list = glob.glob(os.path.join(DP_folder,"*.xlsx"))
AWB_folder = glob.glob(os.path.join(path,"AWB","*.xlsx"))

DP_list_df = pd.read_excel(DP_list[0])

print(DP_list_df.columns)

# 1. Load AWB xlsx file to read
# Empty list for append
AWB = []
# DF merging reference:
# https://saturncloud.io/blog/how-to-merge-multiple-sheets-from-multiple-excel-workbooks-into-a-single-pandas-dataframe/

# Read AWB list in AWB Folder with for loop
for A in AWB_folder:
    # print(A)
    df = pd.read_excel(A)
    AWB.append(df)

# Concat and reset index after appending files
AWB = pd.concat(AWB).reset_index()

# Rename DP List dataframe
# Reference: https://stackoverflow.com/questions/11346283/renaming-column-names-in-pandas
# DP_list_df.rename(columns={'DP No. | Delivery':'DP','Delivery DP':'DP Name'},inplace=True)
DP_list_df = DP_list_df[['DP No. | Delivery', 'Delivery DP']]

New_AWB = pd.merge(AWB, DP_list_df, left_on='DP', right_on='DP No. | Delivery').drop(['DP No. | Delivery','index']
                                                                                     , axis=1)
New_AWB['AWB'] = New_AWB['AWB'].astype(str)

#New_AWB.to_excel(os.path.join(path,"New AWB.xlsx"), index=False)

print(len(AWB))