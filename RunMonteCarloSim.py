# Import libraries
import numpy as np
import pandas as pd
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import yfinance as yf
from scipy.stats import norm
from datetime import date, timedelta

def get_days_until_end_of_year():
    # Get today's date
    today = date.today()

    # Get the last day of the current year
    last_day_of_year = date(today.year, 12, 31)

    # Calculate the difference in days
    days_until_end_of_year = (last_day_of_year - today).days
    return days_until_end_of_year, today.year

def get_last_stock_price(ticker_symbol):
    stock = yf.Ticker(ticker_symbol)
    closing_price = stock.history(period='1d')['Close'].iloc[-1]
    return closing_price

def get_predicted_stock_price_MC (ticker, num_days, num_sims, curr_year):

    if (ticker == "VUSTX"): # error encountered with yfinance for Index funds
        curr_price = 0
    elif (ticker == "VTSMX"): # error encountered with yfinance for Index funds
        curr_price = 0
    elif (ticker == "SWTSX"): # error encountered with yfinance for Index funds
        curr_price = 0
    else:
        curr_price = get_last_stock_price(ticker)

    curr_days = num_days
    curr_sims = num_sims

    # Download Max historical quotes for the given ticker
    data = yf.download(ticker, period="max")
    columns_to_drop = ['Open', 'High', 'Low', 'Adj Close', 'Volume']
    data.drop(columns=columns_to_drop, inplace=True)

    # # Plot Historical Returns
    # plt.figure(figsize=[10,6])
    # title_text = f"Historical Returns - {ticker}"
    # plt.title(title_text)
    # plt.plot(data)
    # plt.show()

    # Historical Log Returns
    log_returns = np.log(1+data.pct_change())

    # # Plot Historical Log Returns
    # plt.figure(figsize=[10,6])
    # title_text = f"Historical Log Returns - {ticker}"
    # plt.title(title_text)
    # plt.plot(log_returns)
    # plt.show()

    # Calculate Drift = [Avg Daily Return - (Variance/2)]
    u = log_returns.mean()
    var = log_returns.var()
    drift = u - (0.5*var)

    # Calculate Std Deviation
    stdev = log_returns.std()

    # Daily Returns
    Z = norm.ppf(np.random.rand(curr_days,curr_sims))
    daily_returns = np.exp(np.array(drift) + np.array(stdev) * Z)

    price_list = np.zeros_like(daily_returns)
    price_list[0] = data.iloc[-1]

    for t in range(1, curr_days):
        price_list[t] = price_list[t-1]*daily_returns[t]
         
    # print(f"Stock: {ticker} ; #Days: {curr_days} Year: {curr_year} (till end of current year); #Simulations: {curr_sims}")
    # print("------------")
    # print("Current Price: ", round(curr_price,2))
    # print("     Expected Mean Price: ", round(np.mean(price_list),2))
    # print("     Quantile (95%): ", round(np.percentile(price_list,95),2))
    # print("     Quantile (5%): ", round(np.percentile(price_list,5),2))

    # # Save to Excel with the actual output of all Simulated values (potential for a large file)
    # df_sim = pd.DataFrame(price_list) 
    # filename = f"{ticker}_MCSims.xlsx"
    # df_sim.to_excel(filename, sheet_name="MCSims",header=False,index=False)

    data = {
        'Num Days' : curr_days,
        'Num Sims' : curr_sims,
        'Curr Price' : [round(curr_price,2)],
        'Year' : curr_year,
        'Quantile (5%)' : [round(np.percentile(price_list,5),2)],
        'Expected Mean Price': [round(np.mean(price_list),2)],
        'Quantile (95%)' : [round(np.percentile(price_list,95),2)]        
    }
    df_new = pd.DataFrame(data)

    # # Add the SUMMARY data to Excel file
    filename = f"{ticker}_MCSims_Summary.xlsx"
    sheet_name = "MCSummary"

    # Load existing Excel file
    try:
        existing_data = pd.read_excel(filename, sheet_name)
        combined_data = pd.concat([existing_data, df_new], ignore_index=True)
    except FileNotFoundError:
        # If file doesn't exist, create a new Excel file
        combined_data = df_new

    # Write combined data back to Excel
    with pd.ExcelWriter(filename, mode='w', engine='openpyxl') as writer:
        combined_data.to_excel(writer, sheet_name=sheet_name, index=False)

    # return price_list

num_simulations = 10000
num_years = 2
stocks = ['VTSMX', 'VUSTX', 'COKE', 'MSFT', 'SWTSX' ]
# stocks = ['VTSMX' ]

# # Iterate over a range of stock tickers
for ticker in stocks:
    num_days, curr_year = get_days_until_end_of_year()
    for j in range(1, num_years+1):  
        get_predicted_stock_price_MC (ticker, num_days, num_simulations, curr_year) 
        num_days = num_days + 365 # Keep adding 365 days (1 Year)
        curr_year = curr_year + 1
