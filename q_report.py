import pandas as pd
import numpy as np
from pandas import read_excel

#读取源数据
file_path = 'files/data.xlsx'
df_fix_assets = read_excel(file_path, sheet_name='固定资产', skiprows=2)
df_employees = read_excel(file_path, sheet_name='人员数量', skiprows=2)
df_budget = read_excel(file_path, sheet_name='预算', skiprows=2)
df_building = read_excel(file_path, sheet_name='建筑面积', skiprows=1)
df_land = read_excel(file_path, sheet_name='土地面积', skiprows=1)
df_business = read_excel(file_path, sheet_name='主营业务', skiprows=1)
df_rent = read_excel(file_path, sheet_name='出租收入', skiprows=2)
#df_rent = read_excel(file_path, sheet_name='出租收入', header=None, names=[['省分', '集团', '集团', '集团', '上市', '上市', '上市'], ['省分', '核算口径', '关联交易', '对外出租', '核算口径', '关联交易', '对外出租']], skiprows=1)
#df =pd.DataFrame()

#表1：预算及完成进度
df_budget_progress = pd.merge(df_budget, df_rent, left_on='省分', right_on='省分')[['省分', '集团预算', '上市预算', '集团-对外出租收入', '上市-对外出租收入']]

df_budget_progress_N10 = df_budget_progress.loc[:9, :]
df_budget_progress_S21 = df_budget_progress.loc[10:, :]
df_budget_progress.loc[31] = ['全国31省', 0, 0, 0, 0]
df_budget_progress.loc[31, 1:5] = df_budget_progress[['集团预算', '上市预算', '集团-对外出租收入', '上市-对外出租收入']].apply(lambda x: x.sum())
df_budget_progress['集团预算进度'] = df_budget_progress['集团-对外出租收入'] / df_budget_progress['集团预算']
df_budget_progress['上市预算进度'] = df_budget_progress['上市-对外出租收入'] / df_budget_progress['上市预算']

df_budget_progress.to_excel('files/output.xlsx')