import pandas as pd
import openpyxl

df = pd.read_excel(r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\NSEBSE_stocks_name.xlsx', engine='openpyxl')

df['REF_SYMBOL'] = df['STOCK_FULL_NAME']

df.sort_values(by = 'STOCK_FULL_NAME', ascending=True, inplace = True)

df.reset_index(drop=True, inplace=True)

df.drop(columns=['Unnamed: 0'], inplace=True, errors='ignore')

rows, columns = df.shape

for e in range((rows)-1):
    if str(df.loc[e,'STOCK_FULL_NAME']) == str(df.loc[e+1,'STOCK_FULL_NAME']) :
        if '.NS' in str(df.loc[e,'STOCK_NAME']) and '.BO' in str(df.loc[e+1,'STOCK_NAME']):
            df.loc[e,'REF_SYMBOL'] = df.loc[e,'STOCK_NAME']
            df.loc[e+1,'REF_SYMBOL'] = df.loc[e,'STOCK_NAME'][:-3]+'.BO'
        elif '.BO' in str(df.loc[e,'STOCK_NAME']) and '.NS' in str(df.loc[e+1,'STOCK_NAME']):
            df.loc[e,'REF_SYMBOL'] = df.loc[e+1,'STOCK_NAME'][:-3]+'.BO'
            df.loc[e+1,'REF_SYMBOL'] = df.loc[e+1,'STOCK_NAME']

for t in range(rows-1):
    if '.NS' in str(df.loc[t,'STOCK_NAME']) :
        df.loc[t,'REF_SYMBOL'] = df.loc[t,'STOCK_NAME']

df.to_excel(r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\NSEBSE_stocks_name.xlsx', engine='openpyxl')