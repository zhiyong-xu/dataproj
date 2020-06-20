import pandas as pd
import numpy as np
from functools import reduce
from pandas import read_excel

#参数定义
quarter = 1


# 读取excel各sheet的源数据
file_path = 'files/data.xlsx'
df_fix_assets = read_excel(file_path, sheet_name='固定资产', skiprows=2)
df_employees = read_excel(file_path, sheet_name='人员数量', skiprows=2)
df_budget = read_excel(file_path, sheet_name='预算', skiprows=2)
df_building = read_excel(file_path, sheet_name='建筑面积', skiprows=1)
df_land = read_excel(file_path, sheet_name='土地面积', skiprows=1)
df_business = read_excel(file_path, sheet_name='主营业务', skiprows=1)
df_rent = read_excel(file_path, sheet_name='出租收入', skiprows=2)
# df_rent = read_excel(file_path, sheet_name='出租收入', header=None, names=[['省分', '集团', '集团', '集团', '上市', '上市', '上市'], ['省分', '核算口径', '关联交易', '对外出租', '核算口径', '关联交易', '对外出租']], skiprows=1)
# df =pd.DataFrame()

# 表1：预算及完成进度原始表
df_budget_progress = pd.merge(df_budget, df_rent, left_on='省分', right_on='省分')[
    ['省分', '集团预算', '上市预算', '集团-对外出租收入', '上市-对外出租收入']]
# 表2：出租收入、出租单价、出租面积原始表
df_rent_area = pd.merge(df_rent, df_building, left_on='省分', right_on='省分')[
    ['省分', '集团-对外出租收入', '建筑总面积', '建筑出租面积']].loc[3:, :].reset_index(drop=True)


# 表5：主营业务收入占用面积原始表
dfs = [df_business, df_building, df_land]
df_revenue_square = reduce(lambda left, right: pd.merge(left, right, on='省分'), dfs)[
    ['省分', '主营业务收入', '建筑总面积', '土地总面积']].loc[3:, :].reset_index(drop=True)

# 表6：固定资产占用面积原始表
dfs = [df_fix_assets, df_building, df_land]
df_assets_square = reduce(lambda left, right: pd.merge(left, right, on='省分'), dfs)[
    ['省分', '净额', '建筑自用面积', '土地自用面积']].loc[3:, :].reset_index(drop=True)


# 表7：利润占用面积原始表
dfs = [df_business, df_building, df_land]
df_profit_square = reduce(lambda left, right: pd.merge(left, right, on='省分'), dfs)[
    ['省分', '利润总额', '建筑总面积', '土地总面积']].loc[3:, :].reset_index(drop=True)

#print(df_profit_square[[df_profit_square.columns[1]]])

# 北10省和南21省预算进度分析
def buget_progress(df_budget_progress=None, region='north10'):
    # 全国31省汇总
    df_budget_progress.loc[31] = ['全国31省', 0, 0, 0, 0]
    df_budget_progress.loc[31, 1:5] = df_budget_progress[['集团预算', '上市预算', '集团-对外出租收入', '上市-对外出租收入']].apply(
        lambda x: x.sum())
    # 分区域分析
    region_start = 0
    region_end = 30
    if region == 'north10':
        region_start = 0
        region_end = 9
        region_count = 10
        region_name = '北10省'
    elif region == 'south21':
        region_start = 10
        region_end = 30
        region_count = 21
        region_name = '南21省'

    df_budget_progress_region = df_budget_progress.copy().loc[region_start:region_end, :]  # 不copy会报错
    df_budget_progress_region.sort_values(by='集团预算', ascending=False, inplace=True)
    df_budget_progress_region = df_budget_progress_region.reset_index(drop=True)
    df_budget_progress_region.loc[region_count] = [region_name, 0, 0, 0, 0]
    df_budget_progress_region.loc[region_count, 1:5] = df_budget_progress_region[['集团预算', '上市预算', '集团-对外出租收入', '上市-对外出租收入']].apply(
        lambda x: x.sum())
    df_budget_progress_region.loc[region_count + 1] = df_budget_progress.loc[31]
    df_budget_progress_region['集团预算进度'] = df_budget_progress_region['集团-对外出租收入'] / df_budget_progress_region['集团预算']
    df_budget_progress_region['上市预算进度'] = df_budget_progress_region['上市-对外出租收入'] / df_budget_progress_region['上市预算']
    # 格式化列
    for i in ['集团预算', '上市预算', '集团-对外出租收入', '上市-对外出租收入']:
        df_budget_progress_region[i] = df_budget_progress_region[i].apply(lambda x: format(round(x, 1), ','))
    for i in ['集团预算进度', '上市预算进度']:
        df_budget_progress_region[i] = df_budget_progress_region[i].apply(lambda x: format(x, '.1%'))
    return df_budget_progress_region

# 表2：出租收入、出租单价、出租面积分析
def rent_area(df_rent_area=None, region='north10'):
    # 全国31省汇总
    df_rent_area.loc[31] = ['全国31省', 0, 0, 0]
    df_rent_area.loc[31, 1:5] = df_rent_area[['集团-对外出租收入', '建筑总面积', '建筑出租面积']].apply(
        lambda x: x.sum())
    # 分区域分析
    region_start = 0
    region_end = 30
    if region == 'north10':
        region_start = 0
        region_end = 9
        region_count = 10
        region_name = '北10省'
    elif region == 'south21':
        region_start = 10
        region_end = 30
        region_count = 21
        region_name = '南21省'

    df_rent_area_region = df_rent_area.copy().loc[region_start:region_end, :]  # 不copy会报错
    df_rent_area_region.sort_values(by='集团-对外出租收入', ascending=False, inplace=True)
    df_rent_area_region.reset_index(drop=True, inplace=True)
    #补充下方区域汇总数据
    df_rent_area_region.loc[region_count] = [region_name, 0, 0, 0]
    df_rent_area_region.loc[region_count, 1:4] = df_rent_area_region[['集团-对外出租收入', '建筑总面积', '建筑出租面积']].apply(
        lambda x: x.sum())
    # 补充下方31省汇总数据
    df_rent_area_region.loc[region_count + 1] = df_rent_area.loc[31]
    # 计算单价和出租率
    df_rent_area_region['出租单价（元/平米/月）'] = df_rent_area_region['集团-对外出租收入'] * 10000 / (df_rent_area_region['建筑出租面积'] * quarter * 3)
    df_rent_area_region['建筑面积出租率'] = df_rent_area_region['建筑出租面积'] / df_rent_area_region['建筑总面积']
    df_rent_area_region['建筑出租面积（万平米）'] = df_rent_area_region['建筑出租面积'] / 10000
    df_rent_area_region = df_rent_area_region[['省分', '集团-对外出租收入', '建筑出租面积（万平米）', '出租单价（元/平米/月）', '建筑面积出租率']]

    # 格式化列
    for i in ['集团-对外出租收入', '建筑出租面积（万平米）', '出租单价（元/平米/月）']:
        df_rent_area_region[i] = df_rent_area_region[i].apply(lambda x: format(round(x, 1), ','))
    for i in ['建筑面积出租率']:
        df_rent_area_region[i] = df_rent_area_region[i].apply(lambda x: format(x, '.1%'))
    return df_rent_area_region

#表头字段列表
table5_names = ['省分', '主营业务收入（百万元）', '建筑面积（万平米）', '收入占用建筑面积（平米/百万元）', '土地面积（万平米）', '收入占用土地面积（平米/百万元）']
table6_names = ['省分', '固定资产金额（百万元）', '建筑自用面积（万平米）', '固定资产占用建筑面积（平米/百万元）', '土地自用面积（万平米）', '固定资产占用土地面积（平米/百万元）']
table7_names = ['省分', '利润总额（百万元）', '建筑面积（万平米）', '利润占用建筑面积（平米/百万元利润）', '土地面积（万平米）', '利润占用土地面积（平米/百万元利润）']

# 收入、固定资产、利润占用面积分析通用函数
def index_area(df=None, region='north10', col_names=table7_names):
    # 全国31省汇总
    df_index_area = df.copy()  #后续会修改列数，先复制一份，否则下一句会出错
    df_index_area.loc[31] = ['全国31省', 0, 0, 0]
    #df_index_area.loc[31, 1:4] = df_index_area[['利润总额', '建筑总面积', '土地总面积']].apply(lambda x: x.sum())
    df_index_area.loc[31, 1:4] = df_index_area[
        [df_index_area.columns[1], df_index_area.columns[2], df_index_area.columns[3]]].apply(
        lambda x: x.sum())
    df_index_area[col_names[3]] = (df_index_area[df_index_area.columns[2]]) / (df_index_area[df_index_area.columns[1]] / 100)
    df_index_area[col_names[5]] = (df_index_area[df_index_area.columns[3]]) / (df_index_area[df_index_area.columns[1]] / 100)
    #print('31省汇总数据-初始时')
    #print(df_index_area.loc[31])

    # 分区域分析
    region_start = 0
    region_end = 30
    if region == 'north10':
        region_start = 0
        region_end = 9
        region_count = 10
        region_name = '北10省'
    elif region == 'south21':
        region_start = 10
        region_end = 30
        region_count = 21
        region_name = '南21省'

    df_index_area_region = df_index_area.copy().loc[region_start:region_end, :]  # 不copy会报错
    #利润为正、为负的分组排序
    df_index_area_pos = df_index_area_region[df_index_area_region[df_index_area_region.columns[1]] > 0].copy()
    df_index_area_neg = df_index_area_region[df_index_area_region[df_index_area_region.columns[1]] < 0].copy()

    df_index_area_pos.sort_values(by=col_names[3], ascending=True, inplace=True)
    df_index_area_neg.sort_values(by=col_names[3], ascending=True, inplace=True)
    df_index_area_all = pd.concat([df_index_area_pos, df_index_area_neg], ignore_index=True)
    df_index_area_all.reset_index(drop=True, inplace=True)
    #补充下方区域汇总数据
    df_index_area_all.loc[region_count] = [region_name, 0, 0, 0, 0, 0]
    df_index_area_all.loc[region_count, 1:4] = df_index_area_all[[df_index_area_all.columns[1], df_index_area_all.columns[2], df_index_area_all.columns[3]]].apply(
        lambda x: x.sum())
    # 补充下方31省汇总数据
    df_index_area_all.loc[region_count + 1] = df_index_area.loc[31]
    # 计算利润占用面积
    df_index_area_all[col_names[3]] = (df_index_area_all[df_index_area_all.columns[2]]) / (df_index_area_all[df_index_area_all.columns[1]] / 100)
    df_index_area_all[col_names[5]] = (df_index_area_all[df_index_area_all.columns[3]]) / (df_index_area_all[df_index_area_all.columns[1]] / 100)

    df_index_area_all[col_names[1]] = df_index_area_all[df_index_area_all.columns[1]] / 100
    df_index_area_all[col_names[2]] = df_index_area_all[df_index_area_all.columns[2]] / 10000
    df_index_area_all[col_names[4]] = df_index_area_all[df_index_area_all.columns[3]] / 10000
    df_index_area_all = df_index_area_all[col_names]

    # 格式化列
    for i in col_names[1:]:
        df_index_area_all[i] = df_index_area_all[i].apply(lambda x: format(round(x, 1), ','))

    return df_index_area_all



df_rent_area_N10 = rent_area(df_rent_area, 'north10')
df_rent_area_S21 = rent_area(df_rent_area, 'south21')

df_budget_progress_N10 = buget_progress(df_budget_progress, 'north10')
df_budget_progress_S21 = buget_progress(df_budget_progress, 'south21')

df_revenue_area_N10 = index_area(df_revenue_square, 'north10', col_names=table5_names)
df_revenue_area_S21 = index_area(df_revenue_square, 'south21', col_names=table5_names)
#print(df_revenue_area_N10)
#print(df_revenue_area_S21)


df_assets_area_N10 = index_area(df_assets_square, 'north10', col_names=table6_names)
df_assets_area_S21 = index_area(df_assets_square, 'south21', col_names=table6_names)

df_profit_area_N10 = index_area(df_profit_square, 'north10', col_names=table7_names)
df_profit_area_S21 = index_area(df_profit_square, 'south21', col_names=table7_names)



with pd.ExcelWriter('files/output.xlsx') as writer:
    df_budget_progress_N10.to_excel(writer, sheet_name='出租收入预算进度', startrow=1)
    df_budget_progress_S21.to_excel(writer, sheet_name='出租收入预算进度', startrow=16)

    df_rent_area_N10.to_excel(writer, sheet_name='出租单价和面积情况', startrow=1)
    df_rent_area_S21.to_excel(writer, sheet_name='出租单价和面积情况', startrow=16)

    df_revenue_area_N10.to_excel(writer, sheet_name='收入占用面积情况', startrow=1)
    df_revenue_area_S21.to_excel(writer, sheet_name='收入占用面积情况', startrow=16)

    df_assets_area_N10.to_excel(writer, sheet_name='固定资产占用面积情况', startrow=1)
    df_assets_area_S21.to_excel(writer, sheet_name='固定资产占用面积情况', startrow=16)

    df_profit_area_N10.to_excel(writer, sheet_name='利润占用面积情况', startrow=1)
    df_profit_area_S21.to_excel(writer, sheet_name='利润占用面积情况', startrow=16)

writer.save()
writer.close()
