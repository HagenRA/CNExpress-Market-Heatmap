
# -*- coding: utf-8 -*-
"""
Created on Tue Jan 21 14:57:59 2020

@author: Administrator
"""

#%% Importing libraries
import pandas as pd
import matplotlib.pyplot as plt
import os
import seaborn as sns

#%% Importing data as df
month_raw = pd.read_excel(r'File Path Here',
					  sheet_name = 'Month')
region_raw = pd.read_excel(r'File Path Here',
					  sheet_name = 'Region')

#%% Check for null values
print(month_raw.isnull().sum())
print(region_raw.isnull().sum())

#%% Cleaning: Drop source link column as not necessary for data analysis
month = month_raw.drop('Source Link', axis=1)

#%% Cleaning: Region data since Spyder can't parse Chinese characters
# Generated the below with a loop code that I deleted because when I was writing this I didn't think of documenting it properly.
region = region_raw
region['Region'].replace('全国','Nationwide', inplace=True)
region['Region'].replace('北京','Beijing', inplace=True)
region['Region'].replace('天津','Tianjin', inplace=True)
region['Region'].replace('河北','Hebei', inplace=True)
region['Region'].replace('山西','Shanxi', inplace=True)
region['Region'].replace('内蒙古','Inner Mongolia', inplace=True)
region['Region'].replace('辽宁','Liaoning', inplace=True)
region['Region'].replace('吉林','Jilin', inplace=True)
region['Region'].replace('黑龙江','Heilongjiang', inplace=True)
region['Region'].replace('上海','Shanghai', inplace=True)
region['Region'].replace('江苏','Jiangsu', inplace=True)
region['Region'].replace('安徽','Anhui', inplace=True)
region['Region'].replace('福建','Fujian', inplace=True)
region['Region'].replace('山东','Shandong', inplace=True)
region['Region'].replace('河南','Henan', inplace=True)
region['Region'].replace('湖北','Hubei', inplace=True)
region['Region'].replace('湖南','Hunan', inplace=True)
region['Region'].replace('广东','Guangdong', inplace=True)
region['Region'].replace('广西','Guangxi', inplace=True)
region['Region'].replace('海南','Hainan', inplace=True)
region['Region'].replace('重庆','Chongqing', inplace=True)
region['Region'].replace('四川','Sichuan', inplace=True)
region['Region'].replace('贵州','Guizhou', inplace=True)
region['Region'].replace('云南','Yunnan', inplace=True)
region['Region'].replace('西藏','Tibet', inplace=True)
region['Region'].replace('陕西','Shaanxi', inplace=True)
region['Region'].replace('甘肃','Gansu', inplace=True)
region['Region'].replace('青海','Qinghai', inplace=True)
region['Region'].replace('宁夏','Ningxia', inplace=True)
region['Region'].replace('新疆','Xinjiang', inplace=True)
region['Region'].replace('浙江','Zhejiang', inplace=True)
region['Region'].replace('江西','Jiangxi', inplace=True)
region = region.rename(columns = {"快递业务量累计（万件）":"Exp Vol (0000 parcels)"})
region = region.rename(columns = {"快递收入累计（万元）":"Exp Revenue (RMB '0000)"})

# Sanity check to make sure all regions converted into English
# Also check to confirm if continuous data is properly represented
#print(region['ID'].unique())
#print(region['Region'].unique())


#%% Visualising monthly growth
month.plot('Month',"Total (0000 parcels)")
plt.xlabel('Month')
plt.ylabel("Total (0000 parcels)")
plt.legend(loc='best')
# Alternatively can be written as:
# month.plot.line(x='Month', y= "Total (0000 parcels)")


#%% Pivot & visualise regional data by volume
region_vol = region.pivot(index = 'ID', columns = 'Region',
							values = '快递业务量累计（万件）')
# NTS normalise data
region_vol.plot()

#%% Pivot & visualise regional data by revenue
region_rev = region.pivot(index = 'ID', columns = 'Region',
							values = '快递收入累计（万元）')

region_rev.plot()

#%%
print(region.isnull().sum())
print(region.shape)

#%%
output = 'File Path Here'
with pd.ExcelWriter(output, engine="openpyxl", mode='a') as writer:
	   region.to_excel(writer, sheet_name = 'Newer_Region')
	   
os.system(f'start "EXCEL.EXE" "{output}" ')

# bokeh / plotly

#%%
#sns.scatterplot(x=None, y=None, hue=None, style=None, size=None, data=None,
#					palette=None, hue_order=None, hue_norm=None, sizes=None,
#					size_order=None, size_norm=None, markers=True,
#					style_order=None, x_bins=None, y_bins=None, units=None,
#					estimator=None, ci=95, n_boot=1000, alpha='auto',
#					x_jitter=None, y_jitter=None, legend='brief', ax=None)

sns.scatterplot(data = region, x = 'ID', y='Exp Vol (0000 parcels)')
