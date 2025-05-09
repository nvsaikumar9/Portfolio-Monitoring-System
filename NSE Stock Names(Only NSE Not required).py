import pandas as pd
from nsetools import Nse
from bsedata.bse import BSE
import openpyxl

nse = Nse()
bse = BSE(update_codes=True)

# Get NSE stock symbols
all_stock_symbols_nse = nse.get_stock_codes()

all_stock_symbols_nse = [i + '.NS' for i in all_stock_symbols_nse]

all_stock_symbols_bse = bse.getScripCodes()

data = {
    'STOCK_NAME': all_stock_symbols_nse #+ [str(code) + '.BO' for code in all_stock_symbols_bse.keys()]
}


# Create a DataFrame and save to Excel
df = pd.DataFrame(data)
df.to_excel(r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\NSE_stocks_name.xlsx', engine='openpyxl')