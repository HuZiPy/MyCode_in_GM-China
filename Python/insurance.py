# -*- coding: utf-8 -*-
"""
Created on Tue Mar 12 10:27:08 2019

@author: KZK0VJ
"""

import pandas as pd
import numpy as np
import os
# from openpyxl import load_workbook

father_content = os.getcwd()
path= father_content + "\\Insurance.xlsx"

df = pd.read_excel(path,sheet_name="Sheet1")

# 选取所需字段
df = df[["model","jv","year","vehicle type","segment","global/local","oem","brand","car/truck","bodystyle_new","oem type",
         "month","volume","domestic/import"]]

# 根据 combine 修改 name 
df.rename(columns={'model': 'Vehicle', 'jv': 'JV', 'year': 'YEAR', 'vehicle type': 'PV/CV', 'segment': 'NEW SEGMENT',
                   "global/local":"GLOBAL/LOCAL","oem":"OEM GROUP","brand":"BRAND","car/truck":"CAR/TRUCK",
                   "oem type":"Car Type","bodystyle_new":"Bodystyle"}, inplace=True) 

# 全改为 大写 
for name in df.columns:
    if name not in ["volume","YEAR"]:
        df[name] = df[name].apply(lambda x: str(x).upper() )
        

# 1. DOMESTIC_CV   
DOMESTIC_CV_P_df = df[(df["domestic/import"]=="DOMESTIC") & ( df["PV/CV"]=="CV" ) & (df["YEAR"]==2019) & (df["OEM GROUP"]!="GM")]
        

DOMESTIC_CV_P_df = DOMESTIC_CV_P_df.fillna("(BLANK)").reset_index(drop=True)

DOMESTIC_CV_df = pd.pivot_table(DOMESTIC_CV_P_df,index=['Vehicle', 'JV', 'YEAR', 'PV/CV', 'NEW SEGMENT', 'GLOBAL/LOCAL','OEM GROUP', 'BRAND', 
                                                        'CAR/TRUCK', 'Bodystyle', 'Car Type'],columns=["month"],values="volume",aggfunc=np.sum).reset_index()

DOMESTIC_CV_df.insert(0,"Source","RETAIL")

DOMESTIC_CV_df.insert(0,"DOMESTIC_CV_df","CATARC_DOMESTIC-CV")


## 2. DOMESTIC_PV_LOCAL
DOMESTIC_PV_LOCAL_P_df = df[(df["domestic/import"]=="DOMESTIC") & ( df["PV/CV"]=="PV" ) & (df['GLOBAL/LOCAL']=='LOCAL') & (df["YEAR"]==2019) & (df["OEM GROUP"]!="GM")]

# fillna 
DOMESTIC_PV_LOCAL_P_df = DOMESTIC_PV_LOCAL_P_df.fillna("(BLANK)").reset_index(drop=True)

# pivot_table
DOMESTIC_PV_LOCAL_df = pd.pivot_table(DOMESTIC_PV_LOCAL_P_df,index=['Vehicle', 'JV', 'YEAR', 'PV/CV', 'NEW SEGMENT', 'GLOBAL/LOCAL','OEM GROUP', 'BRAND', 
                                                                    'CAR/TRUCK', 'Bodystyle', 'Car Type'],columns=["month"],values="volume",aggfunc=np.sum).reset_index()

DOMESTIC_PV_LOCAL_df.insert(0,"Source","RETAIL")

DOMESTIC_PV_LOCAL_df.insert(0,"DOMESTIC_CV_df","CATARC_DOMESTIC-PV-LOCAL")


# 3. IMPORT_CV
IMPORT_P_CV = df[(df["domestic/import"]=="IMPORT") & ( df["PV/CV"]=="CV" ) & (df['GLOBAL/LOCAL']=='GLOBAL') & (df["YEAR"]==2019) & (df["OEM GROUP"]!="GM")]

# fillna 
IMPORT_P_CV = IMPORT_P_CV.fillna("(BLANK)").reset_index(drop=True)

# pivot_table
IMPORT_CV = pd.pivot_table(IMPORT_P_CV,index=['Vehicle', 'JV', 'YEAR', 'PV/CV', 'NEW SEGMENT', 'GLOBAL/LOCAL','OEM GROUP', 'BRAND', 
                                              'CAR/TRUCK', 'Bodystyle', 'Car Type'],columns=["month"],values="volume",aggfunc=np.sum).reset_index()

IMPORT_CV.insert(0,"Source","RETAIL")

IMPORT_CV.insert(0,"DOMESTIC_CV_df","CATARC_IMPORT-CV")


# 4. IMPORT_PV
IMPORT_P_PV = df[(df["domestic/import"]=="IMPORT") & ( df["PV/CV"]=="PV" ) & (df['GLOBAL/LOCAL']=='GLOBAL') & (df["YEAR"]==2019) & (df["OEM GROUP"]!="GM")]

# fillna 
IMPORT_P_PV = IMPORT_P_PV.fillna("(BLANK)").reset_index(drop=True)

# pivot_table
IMPORT_PV = pd.pivot_table(IMPORT_P_PV,index=['Vehicle', 'JV', 'YEAR', 'PV/CV', 'NEW SEGMENT', 'GLOBAL/LOCAL','OEM GROUP', 'BRAND', 
                                              'CAR/TRUCK', 'Bodystyle', 'Car Type'],columns=["month"],values="volume",aggfunc=np.sum).reset_index()

IMPORT_PV.insert(0,"Source","RETAIL")

IMPORT_PV.insert(0,"DOMESTIC_CV_df","CATARC_IMPORT-PV")



# concat
catarc_all = pd.concat([DOMESTIC_CV_df,DOMESTIC_PV_LOCAL_df,IMPORT_CV,IMPORT_PV],ignore_index=True)

catarc_all.to_excel(father_content + "\\catarc.xlsx",index=False)