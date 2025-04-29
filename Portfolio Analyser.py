import openpyxl
import pandas as pd

df_portfolio = pd.read_excel('Portfolio_details.xlsx')

df_API = pd.read_excel('API_stock_prices.xlsx')

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

df1 = pd.DataFrame(New_columns)

df1.to_excel('Portfolio_Analyser.xlsx', index=False, engine='openpyxl')
print(df1)
#-----------------------------------------------------------------------
df = pd.read_excel('LTP.xlsx')

port_ = {
    'PORT_HOLDER' : [],
    'STOCK_NAME' : [],
     'AVG_PRICE' : [],
      'SHARES' : []
      }
'''
class Portfolio():
    def __init__(self, Port_Holder, Name, Avg_price, Lots):
        self.Port_Holder = Port_Holder
        self.Name = Name
        self.Avg_price = Avg_price
        self.Lots = Lots
'''
def Portfolio_details(Port_Holder):
    
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
            port_.update({'PORT_HOLDER' : list1, 'STOCK_NAME' : list2, 'AVG_PRICE' : list3, 'SHARES' : list4} )

while True:
    Port_Holder = input('Enter the name of the account holder...Else enter Done: ')
    if Port_Holder.strip().upper() == 'DONE':
            break
    Port_Holder = Port_Holder.strip().upper()
    Portfolio_details(Port_Holder)
    
df1 = pd.DataFrame(port_)

df1.to_excel("Portfolio_details.xlsx", index= True, engine= 'openpyxl')-