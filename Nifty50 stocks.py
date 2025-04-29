import openpyxl
import pandas as pd

Nifty_stocks = pd.read_csv('MW-NIFTY-50-27-Apr-2025.csv')

#NSE ticket generation block
df = Nifty_stocks

for i in range(len(df)):
    symbol = df.loc[i, 'SYMBOL \n']
    df.loc[i, 'SYMBOL \n'] = symbol + '.NS'

Rows, columns = df.shape
'''
for i in range(1,columns):
    for j in range(len(df)):
        print(i,j)
        value = str(df.iloc[j,i])
        price_value = value.replace(',','')
        df.iloc[j,[i]] = float(price_value)
'''
# print(df.info())
data = df.rename(columns = {'SYMBOL \n' : 'STOCK_NAME'})

data = data['STOCK_NAME']

data.to_excel('NSE_stocks_name.xlsx', index=True, engine='openpyxl')