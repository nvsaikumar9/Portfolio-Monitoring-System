import pandas as pd
from nsetools import Nse

nse = Nse()

all_stock_symbols = nse.get_stock_codes()
all_stock_symbols = [i+'.NS' for i in all_stock_symbols]
data = {
    'STOCK_NAME' : all_stock_symbols
    }

df = pd.DataFrame(data)

df.to_excel('NSE_stocks_name.xlsx', index=True, engine='openpyxl')
