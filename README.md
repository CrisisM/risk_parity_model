# Risk Parity Model Package

This Python package implements a simple risk parity model, which allocates assets in a portfolio such that each asset contributes equally to the portfolio's overall risk. The model calculates the optimal weights for each asset to achieve risk parity, performs backtesting, and visualizes the performance over time.

## Features

- **Risk Parity Strategy Calculation**: This package calculates asset weights based on the risk parity principle, ensuring equal risk contribution from each asset.
- **Covariance Matrix Calculation**: Uses historical data to compute the covariance matrix of asset returns over a configurable time period.
- **Excel Output**: The results, including optimal weights and cumulative returns, are written to an Excel file.
- **Cumulative Returns Visualization**: The package generates a graph showing the cumulative returns of individual assets and the risk parity portfolio.
- **Supports Multiple Asset Classes**: The package can handle multiple asset classes (e.g., domestic and international stocks and bonds).

## Installation

Before running this package, make sure you have the following Python libraries installed:

```bash
pip install pandas numpy matplotlib scipy openpyxl
```

## Usage

### Prepare Input Data:

The input data should be in Excel format with multiple sheets representing different asset classes. Each sheet should contain daily asset prices or returns with the first column being dates.

### Set File Paths:

Modify the `read_file_path` and `write_file_path` variables in the script to point to your input and output Excel files.

### Run the Model:

Call the `Asset_Allocation_risk_parity()` function with your data. The function processes the input, calculates the risk parity weights, and writes the results to the specified Excel file.

Example usage in the script:

```python
Asset_Allocation_risk_parity(dataframes["Domestic Stocks and Bonds"], "Domestic Portfolio", 3)
Asset_Allocation_risk_parity(dataframes["International Stocks and Bonds"], "International Portfolio", 3)
```
