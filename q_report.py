import pandas as pd
import numpy as np
from functools import reduce
from pandas import read_excel

'''
参数定义
'''
# 季度参数，每季度更新
quarter = 2

# 输入输出文件名称设置，每季度更新
in_file_path = 'files/data_20q2.xlsx'
out_file_path = 'files/output_20q2.xlsx'

# 写入文件北10省和南21省的起始行
start_row_dic = {'north10': 1, 'south21': 16}

# 表头字段列表
table1_names = ['省分', '不动产出租收入预算-集团（万元）', '不动产出租收入预算-上市（万元）', '出租收入本年累计-集团（万元）', '出租收入本年累计-上市（万元）', '出租收入进度-集团', '出租收入进度-上市']
table2_names = ['省分', '不动产出租收入本年累计（万元）', '建筑出租面积（万平米）', '出租单价（元/平米/月）', '建筑面积出租率']
table3_names = ['省分', '不动产出租收入本年累计（万元）', '主营业务收入本年累计（万元）', '不动产出租收入与主营业务收入的比例']
table4_names = ['省分', '人数（万人）', '建筑自用面积（万平米）', '人均建筑自用面积（平米/人）', '土地自用面积（万平米）', '人均土地自用面积（平米/人）']
table5_names = ['省分', '主营业务收入（百万元）', '建筑面积（万平米）', '收入占用建筑面积（平米/百万元）', '土地面积（万平米）', '收入占用土地面积（平米/百万元）']
table6_names = ['省分', '固定资产金额（百万元）', '建筑自用面积（万平米）', '固定资产占用建筑面积（平米/百万元）', '土地自用面积（万平米）', '固定资产占用土地面积（平米/百万元）']
table7_names = ['省分', '利润总额（百万元）', '建筑面积（万平米）', '利润占用建筑面积（平米/百万元利润）', '土地面积（万平米）', '利润占用土地面积（平米/百万元利润）']

# 读取excel各sheet
df_fix_assets = read_excel(in_file_path, sheet_name='固定资产', skiprows=2)
df_employees = read_excel(in_file_path, sheet_name='人员数量', skiprows=2)
df_budget = read_excel(in_file_path, sheet_name='预算', skiprows=2)
df_building = read_excel(in_file_path, sheet_name='建筑面积', skiprows=1)
df_land = read_excel(in_file_path, sheet_name='土地面积', skiprows=1)
df_business = read_excel(in_file_path, sheet_name='主营业务', skiprows=1)
df_rent = read_excel(in_file_path, sheet_name='出租收入', skiprows=2)

'''
进行表格关联，生成初始表
'''
# 表1：预算及完成进度原始表
df_budget_progress = pd.merge(df_budget, df_rent, left_on='省分', right_on='省分')[
    ['省分', '集团预算', '上市预算', '集团-对外出租收入', '上市-对外出租收入']]

# 表2：出租收入、出租单价、出租面积原始表
df_rent_area = pd.merge(df_rent, df_building, left_on='省分', right_on='省分')[
    ['省分', '集团-对外出租收入', '建筑总面积', '建筑出租面积']].loc[3:, :].reset_index(drop=True)

# 表3：出租收入和主营业务比例原始表
df_rent_ratio = pd.merge(df_rent, df_business, left_on='省分', right_on='省分')[
    ['省分', '集团-对外出租收入', '主营业务收入']].loc[3:, :].reset_index(drop=True)

# 表4：人均面积原始表
dfs = [df_employees, df_building, df_land]
df_employees_area = reduce(lambda left, right: pd.merge(left, right, on='省分'), dfs)[
    ['省分', '全口径合计', '建筑自用面积', '土地自用面积']].loc[3:, :].reset_index(drop=True)

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


# 北10省和南21省预算进度分析
def budget_progress(df_progress=None, index_name=None, region='north10', col_names=table1_names):
    # 全国31省汇总
    df_progress.loc[31] = ['全国31省', 0, 0, 0, 0]
    df_progress.loc[31, 1:5] = df_progress[[df_progress.columns[1], df_progress.columns[2], df_progress.columns[3], df_progress.columns[4]]].apply(lambda x: x.sum())
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

    df_progress_region = df_progress.copy().loc[region_start:region_end, :]  # 不copy会报错
    df_progress_region.sort_values(by=df_progress_region.columns[1], ascending=False, inplace=True)
    df_progress_region = df_progress_region.reset_index(drop=True)
    df_progress_region.loc[region_count] = [region_name, 0, 0, 0, 0]
    df_progress_region.loc[region_count, 1:5] = df_progress_region[[df_progress_region.columns[1], df_progress_region.columns[2], df_progress_region.columns[3], df_progress_region.columns[4]]].apply(lambda x: x.sum())
    df_progress_region.loc[region_count + 1] = df_progress.loc[31]
    df_progress_region[col_names[5]] = df_progress_region[df_progress_region.columns[3]] / df_progress_region[df_progress_region.columns[1]]
    df_progress_region[col_names[6]] = df_progress_region[df_progress_region.columns[4]] / df_progress_region[df_progress_region.columns[2]]
    # 格式化列
    for i in df_progress_region.columns[1:5]:
        df_progress_region[i] = df_progress_region[i].apply(lambda x: format(round(x, 1), ','))
    for i in df_progress_region.columns[5:7]:
        df_progress_region[i] = df_progress_region[i].apply(lambda x: format(x, '.1%'))
    #更改列名
    df_progress_region.columns = col_names

    return df_progress_region

# 表2：出租收入、出租单价、出租面积分析
def rent_area(df_rent_area=None, index_name=None, region='north10', col_names=table2_names):
    # 全国31省汇总
    df_rent_area.loc[31] = ['全国31省', 0, 0, 0]
    df_rent_area.loc[31, 1:5] = df_rent_area[df_rent_area.columns[1:]].apply(lambda x: x.sum())
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
    df_rent_area_region.sort_values(by=df_rent_area_region.columns[1], ascending=False, inplace=True)
    df_rent_area_region.reset_index(drop=True, inplace=True)
    #补充下方北10省/南21省汇总数据
    df_rent_area_region.loc[region_count] = [region_name, 0, 0, 0]
    df_rent_area_region.loc[region_count, 1:4] = df_rent_area_region[df_rent_area_region.columns[1:4]].apply(
        lambda x: x.sum())
    # 补充下方31省汇总数据
    df_rent_area_region.loc[region_count + 1] = df_rent_area.loc[31]
    # 计算单价和出租率
    df_rent_area_region[col_names[3]] = df_rent_area_region[df_rent_area_region.columns[1]] * 10000 / (df_rent_area_region[df_rent_area_region.columns[3]] * quarter * 3)
    df_rent_area_region[col_names[4]] = df_rent_area_region[df_rent_area_region.columns[3]] / df_rent_area_region[df_rent_area_region.columns[2]]
    #表格显示单位换算，显示为万平米
    df_rent_area_region[df_rent_area_region.columns[3]] = df_rent_area_region[df_rent_area_region.columns[3]] / 10000

    df_rent_area_region.drop(df_rent_area_region.columns[2], axis=1, inplace=True)
    df_rent_area_region.columns = col_names
    #print(df_rent_area_region.columns)

    # 格式化列
    for i in col_names[1:4]:
        df_rent_area_region[i] = df_rent_area_region[i].apply(lambda x: format(round(x, 1), ','))
    for i in col_names[4:]:
        df_rent_area_region[i] = df_rent_area_region[i].apply(lambda x: format(x, '.1%'))

    return df_rent_area_region

# 表3：出租收入与主营业务收入比例分析
def rent_revenue_ratio(df=None, index_name=None, region='north10', col_names=table3_names):
    # 全国31省汇总
    df_ratio = df.copy()  #后续会修改列数，先复制一份，否则下一句会出错
    df_ratio.loc[31] = ['全国31省', 0, 0]
    df_ratio.loc[31, 1:3] = df_ratio[
        [df_ratio.columns[1], df_ratio.columns[2]]].apply(lambda x: x.sum())

    # 增加出租收入占主营业务收入比例列
    df_ratio[col_names[3]] = (df_ratio[df_ratio.columns[1]]) / (df_ratio[df_ratio.columns[2]])

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

    # 北10省和南21省分别计算
    df_ratio_region = df_ratio.copy().loc[region_start:region_end, :]  # 不copy会报错
    # 排序
    df_ratio_region.sort_values(by=col_names[3], ascending=False, inplace=True)
    df_ratio_region.reset_index(drop=True, inplace=True)
    # 北10和南21区域下方追加本区域数据汇总计算行，先不计算主营收入占比的列
    df_ratio_region.loc[region_count] = [region_name, 0, 0, 0]
    df_ratio_region.loc[region_count, 1:3] = df_ratio_region[[df_ratio_region.columns[1], df_ratio_region.columns[2]]].apply(lambda x: x.sum())
    # 追加31省汇总数据行
    df_ratio_region.loc[region_count + 1] = df_ratio.loc[31]
    # 计算主营业务收入占比的列
    df_ratio_region[col_names[3]] = (df_ratio_region[df_ratio_region.columns[1]]) / df_ratio_region[df_ratio_region.columns[2]]
    # 最终表格列重命名
    df_ratio_region.columns = col_names

    # 格式化表格的各列
    for i in col_names[1:3]:
        df_ratio_region[i] = df_ratio_region[i].apply(lambda x: format(round(x, 1), ','))
    for i in col_names[-1:]:
        df_ratio_region[i] = df_ratio_region[i].apply(lambda x: format(x, '.2%'))

    return df_ratio_region


# 表4-表7：人员、收入、固定资产、利润占用面积分析通用函数
def index_area(df=None, index_name=None, region='north10', col_names=table7_names):
    # 全国31省汇总
    df_index_area = df.copy()  #后续会修改列数，先复制一份，否则下一句会出错
    df_index_area.loc[31] = ['全国31省', 0, 0, 0]
    # df_index_area.loc[31, 1:4] = df_index_area[['利润总额', '建筑总面积', '土地总面积']].apply(lambda x: x.sum())
    df_index_area.loc[31, 1:4] = df_index_area[
        [df_index_area.columns[1], df_index_area.columns[2], df_index_area.columns[3]]].apply(
        lambda x: x.sum())

    #分母单位转换
    unit = 1
    if index_name == 'employees':
        unit = 1
    else:
        unit = 100

    # 增加单位指标占用的建筑面积和土地面积两列
    df_index_area[col_names[3]] = (df_index_area[df_index_area.columns[2]]) / (df_index_area[df_index_area.columns[1]] / unit)
    df_index_area[col_names[5]] = (df_index_area[df_index_area.columns[3]]) / (df_index_area[df_index_area.columns[1]] / unit)

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

    # 北10省和南21省分别计算
    df_index_area_region = df_index_area.copy().loc[region_start:region_end, :]  # 不copy会报错
    # 指标为正、为负的分组排序
    df_index_area_pos = df_index_area_region[df_index_area_region[df_index_area_region.columns[1]] > 0].copy()
    df_index_area_neg = df_index_area_region[df_index_area_region[df_index_area_region.columns[1]] < 0].copy()

    df_index_area_pos.sort_values(by=col_names[3], ascending=True, inplace=True)
    df_index_area_neg.sort_values(by=col_names[3], ascending=True, inplace=True)
    df_index_area_all = pd.concat([df_index_area_pos, df_index_area_neg], ignore_index=True)
    df_index_area_all.reset_index(drop=True, inplace=True)
    # 北10和南21区域下方追加本区域数据汇总计算行，先不计算单位指标的列
    df_index_area_all.loc[region_count] = [region_name, 0, 0, 0, 0, 0]
    df_index_area_all.loc[region_count, 1:4] = df_index_area_all[[df_index_area_all.columns[1], df_index_area_all.columns[2], df_index_area_all.columns[3]]].apply(
        lambda x: x.sum())
    # 追加31省汇总数据行
    df_index_area_all.loc[region_count + 1] = df_index_area.loc[31]
    # 统一计算各指标占用面积
    df_index_area_all[col_names[3]] = (df_index_area_all[df_index_area_all.columns[2]]) / (df_index_area_all[df_index_area_all.columns[1]] / unit)
    df_index_area_all[col_names[5]] = (df_index_area_all[df_index_area_all.columns[3]]) / (df_index_area_all[df_index_area_all.columns[1]] / unit)

    # 最终表格显示单位
    index_view_unit = 100
    area_view_unit = 10000
    if index_name in ['employees']:
        index_view_unit = 10000
    df_index_area_all[col_names[1]] = df_index_area_all[df_index_area_all.columns[1]] / index_view_unit
    df_index_area_all[col_names[2]] = df_index_area_all[df_index_area_all.columns[2]] / area_view_unit
    df_index_area_all[col_names[4]] = df_index_area_all[df_index_area_all.columns[3]] / area_view_unit
    df_index_area_all = df_index_area_all[col_names]

    # 格式化表格的各列
    for i in col_names[1:]:
        df_index_area_all[i] = df_index_area_all[i].apply(lambda x: format(round(x, 1), ','))

    return df_index_area_all

# 函数作为参数，构建回调函数，返回dataframe
def get_df(func, df, index, region, columns):
    return func(df, index, region, columns)


if __name__ == '__main__':
    print('*******************************************************')
    print('程序开始运行：')
    # 将函数、dataframe、指标、表头、sheet等做成list，
    func_list = [budget_progress, rent_area, rent_revenue_ratio, index_area, index_area, index_area, index_area]
    df_list = [df_budget_progress, df_rent_area, df_rent_ratio, df_employees_area, df_revenue_square, df_assets_square,
               df_profit_square]
    index_list = [None, None, None, 'employees', 'revenue', 'assets', 'profit']
    table_name_list = [table1_names, table2_names, table3_names, table4_names, table5_names, table6_names, table7_names]
    sheet_name_list = ['出租收入预算进度', '出租单价和面积情况', '出租收入与主营业务收入比例情况', '人均自用面积情况', '收入占用面积情况', '固定资产占用面积情况', '利润占用面积情况']
    # 写入excel的不同sheet
    with pd.ExcelWriter(out_file_path) as writer:
        for i in range(len(func_list)):
            for region in ['north10', 'south21']:  # 北10省和南21省分别写入
                df = get_df(func_list[i], df_list[i], index_list[i], region, table_name_list[i])
                df.to_excel(writer, sheet_name=sheet_name_list[i], startrow=start_row_dic[region])
    writer.save()
    writer.close()
    print('数据已写入excel文件，位置在', out_file_path)
    print('*******************************************************')

