import openpyxl
import yfinance as yf
import pandas as pd

df = pd.read_excel(r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\NSEBSE_stocks_name.xlsx')

port_ = {
    'PORT_HOLDER': [],
    'STOCK_NAME': [],
    'AVG_PRICE': [],
    'SHARES': [],
    'THRESHOLD_LIMIT': [],
    'Email': [],
    'BSE/NSE': [],
    'REF_SYMBOL': []
}

def Portfolio_details(Port_Holder, Thershold_limit, Email):
    
    while True :
        stock_Name = input('Enter the name of the holding stock Eg : TCS, WIPRO...Else enter Done : ')
        stock_Name = stock_Name.strip().upper()

        if stock_Name != 'DONE':
            while True:
                bse_nse = input('Enter the stock exchange name Eg : NSE, BSE... : ')
                if bse_nse.strip().upper() != 'NSE' and bse_nse.strip().upper() != 'BSE':
                    print('Please enter the proper stock exchange name')
                    continue
                elif bse_nse.strip().upper() == 'NSE':
                    stock_Name_ = stock_Name + '.NS'
                    break
                elif bse_nse.strip().upper() == 'BSE':
                    stock_Name_ = stock_Name + '.BO'
                    break

        if stock_Name.strip().upper() == 'DONE':
            break
        elif stock_Name_.strip().upper() not in df['REF_SYMBOL'].values:
            
            print('Stock does not exist, Enter the proper stock name')
            continue

        else:
            Avg_price = float(input('Enter the Average price of stock : '))
            Lots = int(input('Enter number of stocks holding in portfolio : '))

            list1 = port_['PORT_HOLDER']
            list1.append(Port_Holder)

            list2 = port_['STOCK_NAME']
            list2.append(df[df['REF_SYMBOL'] == stock_Name_]['STOCK_NAME'].values[0])
    
            list3 = port_['AVG_PRICE']
            list3.append(Avg_price)
    
            list4 = port_['SHARES']
            list4.append(Lots)   

            list5 = port_['THRESHOLD_LIMIT']
            list5.append(Thershold_limit) 

            list6 = port_['Email']
            list6.append(Email)

            list7 = port_['BSE/NSE']
            list7.append(bse_nse.strip().upper())

            list8 = port_['REF_SYMBOL']
            list8.append(stock_Name)

            port_.update({'PORT_HOLDER' : list1, 'STOCK_NAME' : list2, 'AVG_PRICE' : list3, 'SHARES' : list4, 'THRESHOLD_LIMIT' : list5, 'Email' : list6, 'BSE/NSE' : list7, 'REF_SYMBOL' : list8} )

while True:
    Port_Holder = input('Enter the name of the account holder...Else enter Done: ')
    if Port_Holder.strip().upper() == 'DONE':
            break
    Thershold_limit = float(input('Enter the Threshold limit for a stock : '))
    while True:
        Email = input('Enter the email of the portfolio holder : ')
        if '@' not in Email or '.' not in Email.split('@')[1]:
            print('Please enter a valid email address')
            continue
        else:
            break

    Port_Holder = Port_Holder.strip().upper()
    Portfolio_details(Port_Holder, Thershold_limit, Email)
    
df1 = pd.DataFrame(port_)

df1.to_excel(r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\Portfolio_details.xlsx', index= True, engine= 'openpyxl')