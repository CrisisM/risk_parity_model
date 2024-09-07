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
1. Read Excel Data--------------------------------------------------------------------------------------------------------------------------------
'''

'''
2.1 (Parameters to modify 3) File path settings. You need to extract data from the input Excel file according to the template format and copy the file path to the parameters below.
The pandas openxl function reads the first sheet of the Excel file by default.
'''
# Read the original Excel file
# write_file_path = r"your profile address"
# read_file_path = r"your profile address"

write_file_path = r"your profile address"
read_file_path = r"your profile address"
# Get all sheet names from the Excel file
sheet_names = pd.ExcelFile(read_file_path).sheet_names
# Create an empty dictionary to store DataFrames for different sheets
dataframes = {}
# Read each sheet and store it in the dictionary
for sheet_name in sheet_names:
    # Use pandas' read_excel function to read the sheet data
    df = pd.read_excel(read_file_path, sheet_name=sheet_name)
    # Store the DataFrame in the dictionary with the sheet name as the key
    dataframes[sheet_name] = df


'''
2.2 Output display and content format settings
'''
# Disable scientific notation, retain two decimal places
pd.set_option('display.float_format', lambda x: '%.2f' % x)

# Display all output without limiting the number of rows and columns
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Avoid warnings
warnings.filterwarnings('ignore')

# Check write permissions


def check_write_permission(file_path):
    # Check if the directory has write permissions
    directory = os.path.dirname(file_path)
    if os.access(directory, os.W_OK):
        print(f"Directory {directory} has write permission.")
    else:
        print(f"Directory {directory} does not have write permission.")

    # Check if the file exists and if it has write permissions
    if os.path.exists(file_path):
        if os.access(file_path, os.W_OK):
            print(f"File {file_path} has write permission.")
        else:
            print(f"File {file_path} does not have write permission.")
    else:
        print(f"File {file_path} does not exist.")

# Import stored data


def set_index(close_price_ord):
    close_price_ord['Date'] = pd.to_datetime(
        close_price_ord['Date'], format="%Y%m%d")
    close_price_ord.set_index('Date', inplace=True)


'''
2.3 Result writing module
'''
# Use pandas' ExcelWriter class to create an object for writing Excel files
# Use engine='openpyxl' to specify the openpyxl engine for Excel file writing

# Write results
writer = pd.ExcelWriter(write_file_path, engine='openpyxl',
                        mode='a', if_sheet_exists='replace')
check_write_permission(write_file_path)
# Load the existing Excel file
try:
    book = load_workbook(write_file_path)
except FileNotFoundError:
    # If the file does not exist, an exception will occur, but it's okay, ExcelWriter will automatically create a new file
    book = None


'''
2. Risk Parity Strategy Equations--------------------------------------------------------------------------------------------------------------------------
'''
# Equation 1: Formula for calculating the variance of risk contributions of each asset


def risk_budget_objective(weights, cov):
    weights = np.array(weights)  # weights is a one-dimensional array
    sigma = np.sqrt(np.dot(weights, np.dot(cov, weights)))  # Get portfolio standard deviation
    MRC = np.dot(cov, weights)/sigma  # Calculate marginal risk contribution
    TRC = weights * MRC  # Calculate each asset's contribution to portfolio risk
    delta_TRC = [sum((i - TRC)**2) for i in TRC]  # Calculate the variance of risk contributions between assets
    return sum(delta_TRC)

# Equation 2: Assets cannot be shorted, and the sum of weights equals 1


def total_weight_constraint(x):
    return np.sum(x)-1.0

# Equation 3: Convert decimal to percentage


def percent_formatter(x, pos):
    return f'{x*100:.0f}%'


'''
3. Risk Parity Strategy Calculation Module--------------------------------------------------------------------------------------------------------------------------
'''


def Asset_Allocation_risk_parity(df, sheet_name, frequency):
    '''
    3.1 将净值数据转换为日收益率数据
    '''
    # (1) Get all columns except the first column (date column), these columns contain stock data
    stock_columns = df.columns[1:]
    # (2) Calculate daily returns for each stock and store it in a new DataFrame
    ret = pd.DataFrame()
    # (3) Copy the date column into the new DataFrame
    ret['Date'] = df['Date']
    # (4) Calculate daily returns for each stock and add it to the returns DataFrame
    for stock_column in stock_columns:
        col_name = f'{stock_column}'
        ret[col_name] = df[stock_column].pct_change()  # Use pct_change() to calculate daily returns
    # (5) Remove rows containing NaN values (the first date), resulting in daily return data
    ret = ret.dropna()

    '''
    3.2 Initialize various data arrays
    '''
    # (1) Initial array for monthly returns
    ret_sum = []
    # (2) Initial array for months (x-axis)
    month_x = []
    # (3) Number of asset classes
    k = ret.shape[1] - 1
    # (4) Initialize two-dimensional array for each asset class
    asset_nv = [[] for _ in range(k)]
    # (5) Initialize two-dimensional array for asset allocation percentages
    percen_alloc = [[]for _ in range(k)]

    '''
    3.3 Risk parity strategy calculation for asset weights
    '''
    # 1. Convert the date column to datetime and group the data by month using groupby
    ret['Date'] = pd.to_datetime(ret['Date'])
    monthly_groups = ret.groupby(ret['Date'].dt.to_period('M'))
    # 2. Calculate asset weights for the current month based on the previous n months
    for month, data in monthly_groups:

        # Display data for each month for debugging purposes
        #print(f"Month: {month.to_timestamp().strftime('%Y-%m')}")

        # (1) Calculate data for the previous n months
        previous_data = pd.DataFrame()  # Create an empty DataFrame to store data from the previous n months
        for i in range(frequency - 1, -1, -1):
            previous_month = month - i - 1
            # If data for the previous n months does not exist, skip to the current month
            if previous_month not in monthly_groups.groups:
                break
            previous_data = pd.concat(
                [previous_data, monthly_groups.get_group(previous_month)])

        # (1.1) If the previous n months' data is empty, skip further calculations
        if previous_data.empty:
            continue
        # (1.2) If not empty, proceed with further calculations
        else:
            # (2) Calculate covariance matrix using data from the previous n months
            R_cov = previous_data.cov()
            cov_mon = np.array(R_cov)
            # print(previous_data.head())

            # (3) Calculate allocation for the current month based on last month's data
            # (3.1) Define initial guess values, weights sum to 1
            x0 = np.ones(cov_mon.shape[0]) / cov_mon.shape[0]
            # (3.2) Define boundary conditions
            bnds = tuple((0, None) for x in x0)
            # (3.3) Define constraints, return value must be 0
            cons = ({'type': 'eq', 'fun': total_weight_constraint})
            # (3.4) Iterate multiple times to find the optimal solution (Newton iteration may not have enough iterations)
            options = {'disp': False, 'maxiter': 10000, 'ftol': 1e-20}
            # (3.5) Solve the optimization problem to find the solution that minimizes variance
            solution = minimize(risk_budget_objective, x0, args=(
                cov_mon), bounds=bnds, constraints=cons, options=options)

            # (4) Calculate the returns for each asset during this month
            # (4.1) Select asset columns, assuming asset daily returns data starts from the second column
            asset_returns = data.iloc[:, 1:]
            # (4.2) Calculate daily returns for each asset and get cumulative returns for the month
            cumulative_returns = (1 + asset_returns).cumprod().iloc[-1] - 1
            cumuret = cumulative_returns.values.reshape(1, -1)[0]

            # (5) Calculate the total portfolio return for the month
            retmonth = np.dot(solution.x, cumuret)

            # (6) Store the portfolio return for each month in the array
            ret_sum.append(retmonth)

            # (7) Store the return of each asset for each month in the array
            for i in range(k):
                asset_nv[i].append(cumuret[i])

            # (8) Store the current month as the x-axis value in the array
            mon = month.to_timestamp().strftime('%Y-%m')
            month_x.append(mon)

            # (9) Store the asset allocation percentages for each month in percen_alloc
            for i in range(k):
                percen_alloc[i].append(solution.x[i])

    # 3. Calculate the cumulative portfolio returns for each month and store in ret_sum
    for i in range(0, len(ret_sum)):
        ret_sum[i] = ret_sum[i] + 1
    for i in range(1, len(ret_sum)):
        ret_sum[i] = ret_sum[i-1] * ret_sum[i]

    # 4. Calculate the cumulative returns for each asset for each month and store in asset_nv
    for i in range(k):
        for j in range(0, len(month_x)):
            asset_nv[i][j] = asset_nv[i][j] + 1
        for j in range(1, len(month_x)):
            asset_nv[i][j] = asset_nv[i][j-1] * asset_nv[i][j]

    # 5. Create a DataFrame for the asset allocation percentages for each asset
    df_percen_alloc = pd.DataFrame(percen_alloc).T
    df_month_x = pd.DataFrame(month_x)

    # 6. Merge the month dates with the corresponding data
    merged_df_percen_alloc = pd.concat([df_month_x, df_percen_alloc], axis=1)
    merged_df_percen_alloc.rename(columns=dict(
        zip(merged_df_percen_alloc.columns, ret.columns)), inplace=True)
    merged_df_percen_alloc.columns.values[0] = 'Date'

    # 7. Set the Date column as the index
    merged_df_percen_alloc['Date'] = pd.to_datetime(
        merged_df_percen_alloc['Date'])
    merged_df_percen_alloc.set_index('Date', inplace=True)

    # 8. Use the resample method, 'D' means resampling by day, and backfill each month's value
    merged_df_percen_alloc_resampled = merged_df_percen_alloc.resample(
        'D').ffill()

    # 9. Write the DataFrame to Excel
    merged_df_percen_alloc_resampled.to_excel(
        writer, sheet_name=str(sheet_name), index=True)

    # 10. Save and close the Excel writer
    writer.save()

    # 11. Plot the returns for the risk parity portfolio and individual assets
    # (1) Set the figure size
    plt.figure(figsize=(40, 20))  # Adjust the width and height as needed

    # (2) Plot the returns of the assets and the portfolio
    for i in range(len(asset_nv)):
        plt.plot(month_x, asset_nv[i], label=f'{ret.columns[i]}')
    plt.plot(month_x, ret_sum, label='Risk-Parity Portfolio')

    # (3) Set the x and y labels
    plt.xlabel('Month')
    plt.ylabel('Net Value of Assets and Portfolio')
    plt.title('Net Value of Assets and Portfolio Over Time')

    # (4) Rotate the x-axis labels for better readability
    plt.xticks(rotation=45)

    # (5) Add a legend
    plt.legend(loc='upper left', fontsize='large')

    # (6) Display the plot
    plt.show()



'''
4. Implementation-----------------------------------------------------------------------------------------------------------------------------------
Note: The final number represents the frequency, i.e., how far back in time the covariance matrix is calculated.
'''
Asset_Allocation_risk_parity(dataframes["your"], "your", 3)
Asset_Allocation_risk_parity(dataframes["your"], "your", 3)

writer.close()
