import pandas as pd
from nsetools import Nse
import openpyxl
import yfinance as yf

nse = Nse()

all_stock_symbols = nse.get_stock_codes()
all_stock_symbols = [i+'.NS' for i in all_stock_symbols]
data = {
    'STOCK_NAME' : all_stock_symbols
    }

df1 = pd.DataFrame(data)
df1.to_excel('NSE_stocks_name.xlsx', index=True, engine='openpyxl')

#---------------------------------------------------------------------

# Read stock names from Excel
#df2 = pd.read_excel('NSE_stocks_name.xlsx')
stock_names = df1['STOCK_NAME'].tolist()

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
df2 = pd.DataFrame(LTP)
df2.to_excel('API_stock_prices.xlsx', index=False, engine='openpyxl')

#----------------------------------------------------------------------
'''
#df = pd.read_excel('API_stock_prices.xlsx')
port_ = {
    'PORT_HOLDER': [],
    'STOCK_NAME': [],
    'AVG_PRICE': [],
    'SHARES': [],
    'THRESHOLD_LIMIT': []
}

def Portfolio_details(Port_Holder, Thershold_limit):
    
    while True :
        stock_Name = input('Enter the name of the holding stock Eg : TCS, WIPRO...Else enter Done : ')
        stock_Name = stock_Name.strip().upper()

        if stock_Name.strip().upper() == 'DONE':
            break
        elif stock_Name.strip().upper()+'.NS' not in df2['STOCK_NAME'].values:
            print('Stock does not exist, Enter the proper stock name')
            continue
        else:
            Avg_price = float(input('Enter the Average price of stock : '))
            Lots = int(input('Enter number of stocks holding in portfolio : '))

            list1 = port_['PORT_HOLDER']
            list1.append(Port_Holder)

            list2 = port_['STOCK_NAME']
            list2.append(stock_Name)
    
            list3 = port_['AVG_PRICE']
            list3.append(Avg_price)
    
            list4 = port_['SHARES']
            list4.append(Lots)   

            list5 = port_['THRESHOLD_LIMIT']
            list5.append(Thershold_limit) 
            port_.update({'PORT_HOLDER' : list1, 'STOCK_NAME' : list2, 'AVG_PRICE' : list3, 'SHARES' : list4, 'THRESHOLD_LIMIT' : list5} )

while True:
    Port_Holder = input('Enter the name of the account holder...Else enter Done: ')
    if Port_Holder.strip().upper() == 'DONE':
            break
    Thershold_limit = float(input('Enter the Threshold limit for a stock : '))
    
    Port_Holder = Port_Holder.strip().upper()
    Portfolio_details(Port_Holder, Thershold_limit)
       
df3 = pd.DataFrame(port_)
df3.to_excel("Portfolio_details.xlsx", index= True, engine= 'openpyxl')
#
#----------------------------------------------------------------------
'''

df_portfolio = pd.read_excel('Portfolio_details.xlsx') #3 change
#df_portfolio = df3 #2 change
#df_API = pd.read_excel('API_stock_prices.xlsx')
df_API = df2
Uniq_port_holder = list({df_portfolio['PORT_HOLDER'][i]for i in range(len(df_portfolio['PORT_HOLDER']))})

required_columns_portfolio = ['PORT_HOLDER', 'STOCK_NAME', 'AVG_PRICE', 'SHARES']
required_columns_api = ['STOCK_NAME', 'CLOSE', 'OPEN', 'HIGH', 'LOW']

if not all(col in df_portfolio.columns for col in required_columns_portfolio):
    raise ValueError("Missing required columns in Portfolio_details.xlsx")

if not all(col in df_API.columns for col in required_columns_api):
    raise ValueError("Missing required columns in API_stock_prices.xlsx")

# Initialize variables
Uniq_port_holder = list(df_portfolio['PORT_HOLDER'].unique())
New_columns = {
    'PORT_HOLDER': [], 'STOCK_NAME': [], 'AVG_PRICE': [], 'SHARES': [],
    '%HIGH_CHANGE': [], '%LOW_CHANGE': [], 'CURRENT_%': [], 'HIGH_TO_LOW': [],
    'CLOSE': [], 'OPEN': [], 'HIGH': [], 'LOW': [], 'THRESHOLD_LIMIT': []
}

# Process data
for i in Uniq_port_holder:
    df_port_holder = df_portfolio[df_portfolio['PORT_HOLDER'] == i]
    for j in df_port_holder['STOCK_NAME']:
        # Filter API data
        df_API_stock_values = df_API[df_API['STOCK_NAME'] == j + '.NS']
        if df_API_stock_values.empty:
            print(f"Stock {j+'.NS'} not found in API data. Skipping...")
            continue

        # Extract API values
        close_price = float(df_API_stock_values['CLOSE'].iloc[0])
        open_price = float(df_API_stock_values['OPEN'].iloc[0])
        High_price = float(df_API_stock_values['HIGH'].iloc[0])
        Low_price = float(df_API_stock_values['LOW'].iloc[0])

        # Extract portfolio values
        df_port_stock_vales = df_port_holder[df_port_holder['STOCK_NAME'] == j]
        Average_price = float(df_port_stock_vales['AVG_PRICE'].iloc[0])
        Num_shares = int(df_port_stock_vales['SHARES'].iloc[0])

        # Calculate metrics
        per_change_High = ((High_price - open_price) / open_price) * 100
        per_change_Low = ((open_price - Low_price) / open_price) * 100
        per_currect = ((close_price - open_price) / open_price) * 100
        delta = ((High_price - Low_price) / Low_price) * 100

        # Append to New_columns
        New_columns['PORT_HOLDER'].append(i)
        New_columns['STOCK_NAME'].append(j)
        New_columns['%HIGH_CHANGE'].append(per_change_High)
        New_columns['%LOW_CHANGE'].append(per_change_Low)
        New_columns['CURRENT_%'].append(per_currect)
        New_columns['HIGH_TO_LOW'].append(delta)
        New_columns['AVG_PRICE'].append(Average_price)
        New_columns['SHARES'].append(Num_shares)
        New_columns['CLOSE'].append(close_price)
        New_columns['HIGH'].append(High_price)
        New_columns['LOW'].append(Low_price)
        New_columns['OPEN'].append(open_price)
        New_columns['THRESHOLD_LIMIT'].append(float(df_port_stock_vales['THRESHOLD_LIMIT'].iloc[0]))


# Save to Excel
df4 = pd.DataFrame(New_columns)
output_file = f'Portfolio_Analyser.xlsx'
df4.to_excel(output_file, index=False, engine='openpyxl')


#df = pd.read_excel('Portfolio_Analyser.xlsx')

#current gainer and looser
#the total portfolio value

for j in df4['PORT_HOLDER'].unique():
    # Filter the DataFrame for the current portfolio holder
    df_port_ = df4[df4['PORT_HOLDER']== j]
    rows, columns = df_port_.shape

    Report = ''

    # Calculate the total portfolio value
    total_portfolio_value = [(i*j) for i,j in zip(df_port_['AVG_PRICE'], df_port_['SHARES'])]
    total_portfolio_value = sum(total_portfolio_value)

    current_portfolio_value = [(i*j) for i,j in zip(df_port_['CLOSE'], df_port_['SHARES'])]
    current_portfolio_value = sum(current_portfolio_value)
    profitloss = current_portfolio_value - total_portfolio_value

    '''
    Top_looser_per = df_port_['%HIGH_CHANGE'].min()
    Top_looser_stock = df_port_['STOCK_NAME'][df_port_['%LOW_CHANGE'].idxmin()]
    Top_gainer_stock = df_port_['STOCK_NAME'][df_port_['%HIGH_CHANGE'].idxmax()]
    Top_gainer_per = df_port_['%HIGH_CHANGE'].max()

    Report += f"Top Gainer: {df_port_['STOCK_NAME'][df_port_['CURRENT_%'].idxmax()]} with change of {Top_gainer_per:.2f}%\n"
    Report += f"Top Looser: {df_port_['STOCK_NAME'][df_port_['CURRENT_%'].idxmin()]} with change of {df_port_['%LOW_CHANGE'].min():.2f}%\n"

    '''
    
    Report += f"Portfolio Holder: {j}, P/L : {profitloss:.2f}\n"
    Report += f"Total invested: {total_portfolio_value:.2f}\n"
    Report += f"Current Portfolio Value: {current_portfolio_value:.2f}\n"
    
    for i in range(rows):
        # Extract the stock name and its corresponding values
        stock_name = df_port_['STOCK_NAME'].iloc[i]
        avg_price = df_port_['AVG_PRICE'].iloc[i]
        shares = df_port_['SHARES'].iloc[i]
        threshold_limit = df_port_['THRESHOLD_LIMIT'].iloc[i]
        current_price = df_port_['CLOSE'].iloc[i]
        low_change = df_port_['%LOW_CHANGE'].iloc[i]
        high_change = df_port_['%HIGH_CHANGE'].iloc[i]
        current_price_change = df_port_['CURRENT_%'].iloc[i]

        # Calculate the weightage of the stock in the portfolio
        total_investment_share = avg_price * shares
        weightage = (total_investment_share / total_portfolio_value) * 100

        if threshold_limit < low_change:
            Report += f"Portfolio Holder: {j}, Stock: {stock_name}, crossed threshold and corrected by {low_change:.2f}%, Current of open : {current_price_change:.2f}%, Holding weightage : {weightage:.2f}%\n"
            print(f"Portfolio Holder: {j}, Stock: {stock_name}, crossed threshold by {low_change:.2f}%, Current at : {current_price_change:.2f}%, Holding weightage : {weightage:.2f}%")
        elif threshold_limit < high_change:
            Report += f"Portfolio Holder: {j}, Stock: {stock_name}, crossed threshold and raise by {high_change:.2f}%, Current of price : {current_price_change:.2f}%, Holding weightage : {weightage:.2f}%\n"
            print(f"Portfolio Holder: {j}, Stock: {stock_name}, crossed threshold by {high_change:.2f}%, Current at : {current_price_change:.2f}%, Holding weightage : {weightage:.2f}%")
    

    with open(f'Report {j}.txt', 'w+') as f:
        f.write(Report)
        f.close()