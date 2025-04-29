import openpyxl
import yfinance as yf
import pandas as pd

# Read stock names from Excel
df = pd.read_excel('NSE_stocks_name.xlsx')
stock_names = df['STOCK_NAME'].tolist()

# Initialize LTP dictionary
LTP = {'STOCK_NAME': [], 'CLOSE': [], 'OPEN': [], 'HIGH': [], 'LOW': []}

# Fetch data for all stocks in one API call
price_data = yf.download(tickers=stock_names, period="1d", interval="1d", group_by='ticker', threads=True)

# Process each stock
for stock_name in stock_names:
    if stock_name in price_data:
        stock_data = price_data[stock_name]
        if not stock_data.empty:
            # Extract the latest data
            latest_data = stock_data.iloc[-1]
            LTP['STOCK_NAME'].append(stock_name)
            LTP['CLOSE'].append(latest_data['Close'])
            LTP['OPEN'].append(latest_data['Open'])
            LTP['HIGH'].append(latest_data['High'])
            LTP['LOW'].append(latest_data['Low'])

# Convert LTP dictionary to DataFrame and save to Excel
df1 = pd.DataFrame(LTP)
df1.to_excel('API_stock_prices.xlsx', index=False, engine='openpyxl')

print(df1)