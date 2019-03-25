# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import numpy as np
import os
# from openpyxl import load_workbook

father_content = os.getcwd()
path= father_content + "\\GM's__Sales_Data_of_Vehicle_data.xlsx"

df = pd.read_excel(path)

# pivot_table insurance data only 1-12 all tier
dfNoTier = pd.pivot_table(df,values="volume",index=["brand","model","domestic/import"],columns="month",aggfunc=np.sum)

## convert pivot_table to dataframe 
tt = dfNoTier.reset_index()


## to excel
tt.to_excel( father_content + "\\Tall.xlsx",index=False)


# different tier different excel file
tier_list = ["S6-1 Tier 1 Mix", "S6-2 Tier 2 Mix", "S6-3 Tier 3 Mix", "S6-4 Tier 4 Mix", "S6-5 Tier 5 Mix"]


for tier in tier_list:
    
    df_i = df[df["tier"] == tier[5:11] ]

    dfP_i = pd.pivot_table(df_i,values="volume",index=["brand","model","domestic/import"],columns="month",aggfunc=np.sum)

    dfP_i.reset_index(inplace=True)
    
    dfP_i.to_excel( father_content + "\\" + tier + ".xlsx", index=False)
