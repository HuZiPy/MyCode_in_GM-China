# -*- coding: utf-8 -*-
"""
Created on Tue Feb 26 13:55:15 2019

@author: KZK0VJ
"""

import pandas as pd
import os

opath = os.getcwd()
comb_path = opath + "\\Combined.xlsx"

# 读取 combined data   2019 年的
dfcom = pd.read_excel(comb_path,sheetname="data")
df2019 = dfcom[dfcom["YEAR"]==2019]

df = df2019[["NEW SEGMENT","Vehicle","GLOBAL/LOCAL","RT Jan","RT Feb","RT Mar","RT Apr","RT May","RT Jun","RT Jul", "RT Aug","RT Sep","RT Oct","RT Nov","RT Dec","CYTD"]]

last_df = df.reset_index(drop=True)


## by vehicle type

file_path = opath + "\\"
items = ["Car-A","Car-B","MPV-B","MPV-C","Pickup-B","SUV-B","SUV-C","Van-B","Van-D"]

for item in items:
    
    if item in ['SUV-B','SUV-C']:
        
        dfs = last_df[(last_df["NEW SEGMENT"]=='SUV-B') & (last_df["GLOBAL/LOCAL"]=='LOCAL')]

        dfs = dfs.reset_index(drop=True)
        
        tts = dfs[[ 'CYTD','RT Jan', 'RT Feb', 'RT Mar', 'RT Apr', 'RT May', 'RT Jun','RT Jul', 'RT Aug', 'RT Sep', 'RT Oct', 'RT Nov', 'RT Dec']].groupby(dfs['Vehicle']).sum()

        dfsuv = tts.sort_values(by="CYTD",ascending=False).iloc[:10,:]
        
        dfsuv.to_excel(file_path + item + ".xlsx")
        
        
    else:
        
        dfi = last_df[last_df["NEW SEGMENT"]==item]

        dfi = dfi.reset_index(drop=True)

        tt = dfi[[ 'CYTD','RT Jan', 'RT Feb', 'RT Mar', 'RT Apr', 'RT May', 'RT Jun','RT Jul', 'RT Aug', 'RT Sep', 'RT Oct', 'RT Nov', 'RT Dec']].groupby(dfi['Vehicle']).sum()

        dfitem = tt.sort_values(by="CYTD",ascending=False).iloc[:10,:]

        dfitem.to_excel(file_path + item + ".xlsx")
        
        
## by OEM GROUP

dfoem = df2019[["OEM GROUP","CYTD","RT Jan","RT Feb","RT Mar","RT Apr","RT May","RT Jun","RT Jul", "RT Aug","RT Sep","RT Oct","RT Nov","RT Dec"]]

dfo = dfoem.reset_index(drop=True)

ttoem = dfo[[ 'CYTD','RT Jan', 'RT Feb', 'RT Mar', 'RT Apr', 'RT May', 'RT Jun','RT Jul', 'RT Aug', 'RT Sep', 'RT Oct', 'RT Nov', 'RT Dec']].groupby(dfo["OEM GROUP"]).sum()

df_OEM = ttoem.sort_values(by="CYTD",ascending=False).iloc[:20,:]

df_OEM.to_excel(file_path+"oem.xlsx")