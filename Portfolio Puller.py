import openpyxl
import yfinance as yf
import pandas as pd

df = pd.read_excel('API_stock_prices.xlsx')

port_ = {
    'PORT_HOLDER': [],
    'STOCK_NAME': [],
    'AVG_PRICE': [],
    'SHARES': [],
    'THRESHOLD_LIMIT': [],
    'Email': []
}

def Portfolio_details(Port_Holder, Thershold_limit, Email):
    
    while True :
        stock_Name = input('Enter the name of the holding stock Eg : TCS, WIPRO...Else enter Done : ')
        stock_Name = stock_Name.strip().upper()

        if stock_Name.strip().upper() == 'DONE':
            break
        elif stock_Name.strip().upper()+'.NS' not in df['STOCK_NAME'].values:
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

            list6 = port_['Email']
            list6.append(Email)
            port_.update({'PORT_HOLDER' : list1, 'STOCK_NAME' : list2, 'AVG_PRICE' : list3, 'SHARES' : list4, 'THRESHOLD_LIMIT' : list5, 'Email' : list6} )

while True:
    Port_Holder = input('Enter the name of the account holder...Else enter Done: ')
    if Port_Holder.strip().upper() == 'DONE':
            break
    Thershold_limit = float(input('Enter the Threshold limit for a stock : '))
    Email = input('Enter the email of the portfolio holder : ')

    Port_Holder = Port_Holder.strip().upper()
    Portfolio_details(Port_Holder, Thershold_limit, Email)
    
df1 = pd.DataFrame(port_)

df1.to_excel("Portfolio_details.xlsx", index= True, engine= 'openpyxl')