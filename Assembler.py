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

#df = pd.read_excel('API_stock_prices.xlsx')

port_ = {
    'PORT_HOLDER' : [],
    'STOCK_NAME' : [],
     'AVG_PRICE' : [],
      'SHARES' : []
      }
def Portfolio_details(Port_Holder):
    
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
            port_.update({'PORT_HOLDER' : list1, 'STOCK_NAME' : list2, 'AVG_PRICE' : list3, 'SHARES' : list4} )

while True:
    Port_Holder = input('Enter the name of the account holder...Else enter Done: ')
    if Port_Holder.strip().upper() == 'DONE':
            break
    Port_Holder = Port_Holder.strip().upper()
    Portfolio_details(Port_Holder)
    
df3 = pd.DataFrame(port_)
df3.to_excel("Portfolio_details.xlsx", index= True, engine= 'openpyxl')

#----------------------------------------------------------------------

#df_portfolio = pd.read_excel('Portfolio_details.xlsx')
df_portfolio = df3
#df_API = pd.read_excel('API_stock_prices.xlsx')
df_API = df2
Uniq_port_holder = list({df_portfolio['PORT_HOLDER'][i]for i in range(len(df_portfolio['PORT_HOLDER']))})

New_columns = {'PORT_HOLDER': [] ,'STOCK_NAME' : [], 'AVG_PRICE': [],'SHARES': [], '%HIGH_CHANGE' : [], '%LOW_CHANGE' : [], 'CURRENT_%' : [], 'HIGH_TO_LOW' : []}

for i in Uniq_port_holder:
    df_port_holder = df_portfolio[df_portfolio['PORT_HOLDER'] == i]
    for j in df_port_holder['STOCK_NAME']:

        #values of API
        df_API_stock_values = df_API[df_API['STOCK_NAME'] == j+'.NS']
        close_price = float(df_API_stock_values['CLOSE'].iloc[0])
        open_price = float(df_API_stock_values['OPEN'].iloc[0])
        High_price = float(df_API_stock_values['HIGH'].iloc[0])
        Low_price = float(df_API_stock_values['LOW'].iloc[0])

        #Values of portfolio
        df_port_stock_vales = df_port_holder[df_port_holder['STOCK_NAME']==j]
        Average_price = float(df_port_stock_vales['AVG_PRICE'].iloc[0])
        Num_shares = int(df_port_stock_vales['SHARES'].iloc[0])

        per_change_High = ((High_price - open_price)/open_price)*100
        per_change_Low = ((open_price - Low_price)/open_price)*100
        per_currect = ((close_price - open_price)/open_price)*100
        delta = ((High_price - Low_price)/Low_price)*100

        New_columns['PORT_HOLDER'].append(i)
        New_columns['STOCK_NAME'].append(j)
        New_columns['%HIGH_CHANGE'].append(per_change_High)
        New_columns['%LOW_CHANGE'].append(per_change_Low)
        New_columns['CURRENT_%'].append(per_currect)
        New_columns['HIGH_TO_LOW'].append(delta)
        New_columns['AVG_PRICE'].append(Average_price)
        New_columns['SHARES'].append(Num_shares)

df4 = pd.DataFrame(New_columns)
df4.to_excel('Portfolio_Analyser.xlsx', index=False, engine='openpyxl')

