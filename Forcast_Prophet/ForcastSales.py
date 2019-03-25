# -*- coding: utf-8 -*-
"""
Created on Mon Mar 25 14:49:05 2019

@author: Hu Wei
"""

import pandas as pd
from fbprophet import Prophet


df = pd.read_excel("filename",sheet_name="retail")

# for example CADILLAC
tdf = df[["ds","CADILLAC"]]

tdf.rename(columns={"CADILLAC":"y"},inplace=True)

m = Prophet()
m.add_country_holidays(country_name="China")

m.fit(tdf)

future = m.make_future_dataframe(periods=10,freq='M')

forecast = m.predict(future)

# plot
fig1 = m.plot(forecast)

fig = m.plot_components(forecast)