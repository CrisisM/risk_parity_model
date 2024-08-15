# -*- coding: utf-8 -*-
"""
Created on Mon Oct 30 22:21:14 2023

@author: Wenbo Liu

"""

import os
import pandas as pd
import numpy as np
import warnings
from openpyxl import load_workbook
from scipy.optimize import minimize
import matplotlib.pyplot as plt


'''
一、读取Excel数据--------------------------------------------------------------------------------------------------------------------------------
'''


'''
2.1（需要修改的参数3）文件路径设置，需要在输入的excel里按模版格式提取数据，将文件路径复制给下面的参数
pandas的openxsl函数默认读取excel的第一张工作表
'''
# 读取原始Excel文件
# write_file_path = r"F:\A2 public共享区\刘文博-202310\a3.模型策略\a6.风险平价封装\风险平价策略输出.xlsx"
# read_file_path = r"F:\A2 public共享区\刘文博-202310\a3.模型策略\a6.风险平价封装\风险平价策略输入.xlsx"

write_file_path = r"C:\Users\刘\Nutstore\1\A2 public共享区\刘文博-202310\a4.模型策略\a1.风险平价模型封装\风险平价策略输出.xlsx"
read_file_path = r"C:\Users\刘\Nutstore\1\A2 public共享区\刘文博-202310\a4.模型策略\a1.风险平价模型封装\风险平价策略输入.xlsx"
# 获取Excel文件中所有工作表的名称
sheet_names = pd.ExcelFile(read_file_path).sheet_names
# 创建一个空字典，用于存储不同工作表的DataFrame
dataframes = {}
# 依次读取每个工作表并存储在字典中
for sheet_name in sheet_names:
    # 使用pandas的read_excel函数读取工作表数据
    df = pd.read_excel(read_file_path, sheet_name=sheet_name)
    # 将DataFrame存储在字典中，以工作表名称作为键
    dataframes[sheet_name] = df


'''
2.2 输出显示与内容格式设置
'''
# 禁止使用科学计数法，保留2个小数位
pd.set_option('display.float_format', lambda x: '%.2f' % x)

# 全部输出显示，不限制输出显示行列数（不设置的话默认20）
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# 避免warning
warnings.filterwarnings('ignore')

# 检查写入权限


def check_write_permission(file_path):
    # 检查目录是否具有写入权限
    directory = os.path.dirname(file_path)
    if os.access(directory, os.W_OK):
        print(f"目录 {directory} 具有写入权限.")
    else:
        print(f"目录 {directory} 没有写入权限.")

    # 检查文件是否存在，并检查文件是否具有写入权限
    if os.path.exists(file_path):
        if os.access(file_path, os.W_OK):
            print(f"文件 {file_path} 具有写入权限.")
        else:
            print(f"文件 {file_path} 没有写入权限.")
    else:
        print(f"文件 {file_path} 不存在.")

# 导入存储的数据


def set_index(close_price_ord):
    close_price_ord['Date'] = pd.to_datetime(
        close_price_ord['Date'], format="%Y%m%d")
    close_price_ord.set_index('Date', inplace=True)


'''
2.3 结果写入模块
'''
# 使用pandas库的ExcelWriter类来创建一个用于写入Excel文件的对象writer
# 使用engine='openpyxl'参数，表示使用openpyxl作为底层引擎来进行Excel文件的写入操作

# 结果写入
writer = pd.ExcelWriter(write_file_path, engine='openpyxl',
                        mode='a', if_sheet_exists='replace')
check_write_permission(write_file_path)
# 加载已存在的Excel文件
try:
    book = load_workbook(write_file_path)
except FileNotFoundError:
    # 如果文件不存在，会出现异常，不过没关系，ExcelWriter会自动创建新文件
    book = None


'''
二、风险平价策略相关方程--------------------------------------------------------------------------------------------------------------------------
'''
# 方程1：计算各个资产风险贡献的方差的方程


def risk_budget_objective(weights, cov):
    weights = np.array(weights)  # weights为一维数组
    sigma = np.sqrt(np.dot(weights, np.dot(cov, weights)))  # 获取组合标准差
    #sigma = np.sqrt(weights@cov@weights)
    MRC = np.dot(cov, weights)/sigma  # 计算边际风险贡献
    TRC = weights * MRC  # 计算单个资产对投资组合的风险贡献
    delta_TRC = [sum((i - TRC)**2) for i in TRC]  # 计算各个资产之间风险贡献的方差
    return sum(delta_TRC)

# 方程2：资产不可以做空，权重和为1


def total_weight_constraint(x):
    return np.sum(x)-1.0

# 方程3：将小数转换为百分比


def percent_formatter(x, pos):
    return f'{x*100:.0f}%'


'''
三、风险平价策略计算模块--------------------------------------------------------------------------------------------------------------------------
'''


def Asset_Allocation_risk_parity(df, sheet_name, frequency):
    '''
    3.1 将净值数据转换为日收益率数据
    '''
    # (1)获取除了第一列（日期列）之外的所有列，这些列包含股票数据
    stock_columns = df.columns[1:]
    # (2)计算每只股票的日收益率并存储为新的DataFrame
    ret = pd.DataFrame()
    # (3)将日期列复制到新的DataFrame中
    ret['Date'] = df['Date']
    # (4)计算每只股票的日收益率并添加到returns_df中
    for stock_column in stock_columns:
        col_name = f'{stock_column}'
        ret[col_name] = df[stock_column].pct_change()  # 使用pct_change()计算日收益率
    # (5)删除包含NaN值的行（第一个日期的数据）,得到日收益率数据
    ret = ret.dropna()

    '''
    3.2 初始化各类型数据数组
    '''
    # (1)初始纵坐标每个月收益率的数组
    ret_sum = []
    # (2)初始横坐标月的数组
    month_x = []
    # (3)求大类资产有多少个
    k = ret.shape[1] - 1
    # (4)初始化各个大类资产的二维数组
    asset_nv = [[] for _ in range(k)]
    # (5)初始化大类资产比例的二维数组
    percen_alloc = [[]for _ in range(k)]

    '''
    3.3 风险平价策略计算各个资产的权重
    '''
    # 1.将日期列转换为日期时间类型并使用groupby按月分组数据
    ret['Date'] = pd.to_datetime(ret['Date'])
    monthly_groups = ret.groupby(ret['Date'].dt.to_period('M'))
    # 2.按前n个月计算当月资产的权重
    for month, data in monthly_groups:

        # 显示每个月的数据，debug时方便观测
        #print(f"Month: {month.to_timestamp().strftime('%Y-%m')}")

        # (1)计算前n个月的数据
        previous_data = pd.DataFrame()  # 创建一个空的DataFrame用于存储前n个月的数据
        for i in range(frequency - 1, -1, -1):
            previous_month = month - i - 1
            # 如果前n个月的数据不存在就不得到当前月的数据
            if previous_month not in monthly_groups.groups:
                break
            previous_data = pd.concat(
                [previous_data, monthly_groups.get_group(previous_month)])

        # (1.1)如果前面n个月的数据为空则不进行后面的运算
        if previous_data.empty:
            continue
        # (1.2)不为空则继续后面的运算
        else:
            # (2)使用前n个月的数据计算协方差矩阵
            R_cov = previous_data.cov()
            cov_mon = np.array(R_cov)
            # print(previous_data.head())

            # (3)使用上个月的数据计算当月的配比
            # (3.1)定义初始猜测值,权重和为1
            x0 = np.ones(cov_mon.shape[0]) / cov_mon.shape[0]
            # (3.2)定义边界条件
            bnds = tuple((0, None) for x in x0)
            # (3.3)定义约束条件，返回值为0
            cons = ({'type': 'eq', 'fun': total_weight_constraint})
            # (3.4)多次迭代求最优解（牛顿迭代可能迭代次数过少）
            options = {'disp': False, 'maxiter': 10000, 'ftol': 1e-20}
            # (3.5)求最优化问题：方差最小值时权重的解
            solution = minimize(risk_budget_objective, x0, args=(
                cov_mon), bounds=bnds, constraints=cons, options=options)

            # (4)计算这个月每个标的的收益率
            # (4.1)选择大类资产列，假设大类资产的日收益率数据从第二列开始
            asset_returns = data.iloc[:, 1:]
            # (4.2)计算每个大类资产的每天收益率，并计算累积乘积得到每个月各资产的收益率
            cumulative_returns = (1 + asset_returns).cumprod().iloc[-1] - 1
            cumuret = cumulative_returns.values.reshape(1, -1)[0]

            # (5)计算这个月的资产组合总收益率
            retmonth = np.dot(solution.x, cumuret)

            # (6)将每个月的资产组合收益率存入数组
            ret_sum.append(retmonth)

            # (7)将每个月各个资产的收益率存入数组
            for i in range(k):
                asset_nv[i].append(cumuret[i])

            # (8)将当月作为横坐标存入数组
            mon = month.to_timestamp().strftime('%Y-%m')
            month_x.append(mon)

            # (9)将每月的资产比例存入percen_alloc
            for i in range(k):
                percen_alloc[i].append(solution.x[i])

    # 3.求每个月的投资组合累计收益率并存入ret_sum
    for i in range(0, len(ret_sum)):
        ret_sum[i] = ret_sum[i] + 1
    for i in range(1, len(ret_sum)):
        ret_sum[i] = ret_sum[i-1] * ret_sum[i]

    # 4.求每个月的各个资产的累计收益率并存入asset_nv
    for i in range(k):
        for j in range(0, len(month_x)):
            asset_nv[i][j] = asset_nv[i][j] + 1
        for j in range(1, len(month_x)):
            asset_nv[i][j] = asset_nv[i][j-1] * asset_nv[i][j]

    # 5.创建各个资产投资组合比例的Dataframe
    df_percen_alloc = pd.DataFrame(percen_alloc).T
    df_month_x = pd.DataFrame(month_x)

    # 6.合并月日期和相对应的数据
    merged_df_percen_alloc = pd.concat([df_month_x, df_percen_alloc], axis=1)
    merged_df_percen_alloc.rename(columns=dict(
        zip(merged_df_percen_alloc.columns, ret.columns)), inplace=True)
    merged_df_percen_alloc.columns.values[0] = 'Date'

    # 7.将Date列设置为索引
    merged_df_percen_alloc['Date'] = pd.to_datetime(
        merged_df_percen_alloc['Date'])
    merged_df_percen_alloc.set_index('Date', inplace=True)

    # 8.使用resample方法，D表示按天重新采样,并向后填充每个月的值
    merged_df_percen_alloc_resampled = merged_df_percen_alloc.resample(
        'D').ffill()

    # 9.将dataframe写入excel
    merged_df_percen_alloc_resampled.to_excel(
        writer, sheet_name=str(sheet_name), index=True)

    # 10.保存并关闭excel
    writer.save()

    # 11.画出资产风险平价投资组合和资产的收益率曲线
    # (1)设置图像的大小
    plt.figure(figsize=(40, 20))  # Adjust the width and height as needed

    # (2)画出资产和投资组合的收益率曲线
    for i in range(len(asset_nv)):
        plt.plot(month_x, asset_nv[i], label=f'{ret.columns[i]}')
    plt.plot(month_x, ret_sum, label='Risk-Parity Portfolio')

    # (3)设置横纵坐标
    plt.xlabel('Month')
    plt.ylabel('Net Value of Assets and Portfolio')
    plt.title('Net Value of Assets and Portfolio Over Time')

    # (4)使横坐标看得更清晰
    plt.xticks(rotation=45)

    # (5)加标记
    plt.legend(loc='upper left', fontsize='large')

    # (6)画出图像
    plt.show()


'''
四、实现-----------------------------------------------------------------------------------------------------------------------------------
注：最后数字代表frequency，即往前多久的时间计算协方差矩阵
'''
Asset_Allocation_risk_parity(dataframes["国内股债商"], "国内股债商", 3)
Asset_Allocation_risk_parity(dataframes["国外股债商"], "国外股债商", 3)

writer.close()
