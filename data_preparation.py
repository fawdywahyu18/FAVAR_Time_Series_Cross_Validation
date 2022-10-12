# -*- coding: utf-8 -*-
"""
Created on Wed Jun  8 09:18:04 2022

@author: Fawdy
"""

import pandas as pd
import openpyxl
from pandas import ExcelWriter

main_wd = 'E:/Kerja/UbuntuFiles'
backup_wd = 'G:/Backup Fawdy/Kerja/UbuntuFiles'
wd = backup_wd
source_file = wd+'/indeks_harga_konsumen updated 2022-10-05.xlsx'
wb = openpyxl.load_workbook(source_file)
sheet_names = wb.sheetnames

monthly_sheets = sheet_names[4:10]
daily_sheets = sheet_names[10:12]
annualy_sheets = [sheet_names[-1]]

# Monthly Series
cpi = pd.read_excel(source_file, sheet_name=sheet_names[3])
cpi['date_index'] = pd.date_range(start='1960-01-01', periods=len(cpi), freq='MS')
cpi.set_index('date_index', inplace=True)
cpi = cpi.drop(columns=['Date'])

concat_df = cpi
for i in monthly_sheets:
    test = pd.read_excel(source_file, sheet_name=i)
    column_names = list(test)
    test.rename(columns = {column_names[0]:'date_index'}, inplace=True)
    test = test.loc[25:,:]
    start_date = test['date_index'][25]
    test['date_index'] = pd.date_range(start=start_date, periods=len(test), freq='MS')
    test.set_index('date_index', inplace=True)
    concat_df = pd.concat([concat_df, test], axis=1)

# Daily Data
concat_daily = pd.DataFrame()
for i in daily_sheets:
    d_df = pd.read_excel(source_file, sheet_name=i)
    colnames = list(d_df)
    d_df = d_df.loc[25:]
    no_na = d_df.loc[d_df[colnames[1]].dropna().index]
    no_na.set_index('Unnamed: 0', inplace=True)
    concat_daily = pd.concat([concat_daily, no_na], axis=1)


concat_df = pd.concat([concat_df, concat_daily], axis=1)
concat_df = concat_df.loc[concat_df.index>='2005-01-01']

# Annual Data
a_df = pd.read_excel(source_file, sheet_name=annualy_sheets[0])
a_colnames = list(a_df)[1:]
a_df = a_df.loc[25:,:]
a_df['date_index'] = pd.date_range(start='1991-01-01', 
                                   periods=len(a_df), freq='AS')
a_df = a_df.drop(columns=['Unnamed: 0'])
a_df.set_index('date_index', inplace=True)
a_to_m = pd.DataFrame()
for col in a_colnames:
    a_to_m[col] = a_df[col].resample('MS').interpolate(limit=36, limit_direction='forward')

monthly_df = pd.concat([concat_df, a_to_m], axis=1)
monthly_df = monthly_df.loc[(monthly_df.index>='2005-01-01') & (monthly_df.index<='2022-09-01')] # diganti per tanggal update terakhir
monthly_df.to_excel(wd+'/ihk_prepared new.xlsx', sheet_name='multivariate_monthly')









