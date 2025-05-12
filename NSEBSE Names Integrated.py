import pandas as pd
from nsetools import Nse
from bsedata.bse import BSE
import openpyxl
import yfinance as yf
import time
import os
from pathlib import Path

nse = Nse()
bse = BSE(update_codes=True)

# Get NSE stock symbols
Raw_stock_symbols_nse = nse.get_stock_codes()

All_stock_symbols_nse = [i.strip().upper() + '.NS' for i in Raw_stock_symbols_nse]

All_stock_names_nse = []
for i in All_stock_symbols_nse:
    try:
        # Fetch the stock name using yfinance
        ticker = yf.Ticker(i)
        name = ticker.info.get("longName", "Name not found")
        nse_name = name.upper().replace("LIMITED", "").replace("THE","").replace("LTD", "").replace("LTD.", "").replace("PVT", "").replace("PVT.", "").replace("PRIVATE", "").replace("PRIVATE.", "").replace("CORPORATION", "").replace("CORP", "").replace("CORP.", "").replace("INDUSTRIES", "").replace("INDUSTRIES.", "").replace(" .","").strip()
        All_stock_names_nse.append(nse_name)
    except Exception as e:
        name = ""
        All_stock_names_nse.append(name)
        print(f"Error fetching data for {i}: {e}")
        

all_stock_symbols_bse = bse.getScripCodes()
bse_name = all_stock_symbols_bse.values()
bse_name = [bname.upper().replace("LIMITED", "").replace("THE","").replace("LTD", "").replace("LTD.", "").replace("PVT", "").replace("PVT.", "").replace("PRIVATE", "").replace("PRIVATE.", "").replace("CORPORATION", "").replace("CORP", "").replace("CORP.", "").replace("INDUSTRIES", "").replace("INDUSTRIES.", "").replace(" .","").strip() for bname in bse_name]

data = {
    'STOCK_NAME': All_stock_symbols_nse + [str(code) + '.BO' for code in all_stock_symbols_bse.keys()],
    'STOCK_FULL_NAME': All_stock_names_nse + [values.strip().upper() for values in bse_name]
}

# Create a DataFrame and save to Excel
df = pd.DataFrame(data)
#df.to_excel(r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\NSEBSE_stocks_namebackupdata.xlsx', engine='openpyxl')

df['REF_SYMBOL'] = df['STOCK_FULL_NAME']

df.sort_values(by = 'STOCK_FULL_NAME', ascending=True, inplace = True)

df.reset_index(drop=True, inplace=True)

df.drop(columns=['Unnamed: 0'], inplace=True, errors='ignore')

rows, columns = df.shape

for e in range((rows)-1):
    if str(df.loc[e,'STOCK_FULL_NAME']) == str(df.loc[e+1,'STOCK_FULL_NAME']) :
        if '.NS' in str(df.loc[e,'STOCK_NAME']) and '.BO' in str(df.loc[e+1,'STOCK_NAME']):
            df.loc[e+1,'STOCK_NAME'] = df.loc[e,'STOCK_NAME'][:-3]+'.BO'
            df.loc[e,'REF_SYMBOL'] = df.loc[e,'STOCK_NAME']
            df.loc[e+1,'REF_SYMBOL'] = df.loc[e,'STOCK_NAME'][:-3]+'.BO'
        elif '.BO' in str(df.loc[e,'STOCK_NAME']) and '.NS' in str(df.loc[e+1,'STOCK_NAME']):
            df.loc[e,'STOCK_NAME'] = df.loc[e+1,'STOCK_NAME'][:-3]+'.BO'
            df.loc[e,'REF_SYMBOL'] = df.loc[e+1,'STOCK_NAME'][:-3]+'.BO'
            df.loc[e+1,'REF_SYMBOL'] = df.loc[e+1,'STOCK_NAME']

for t in range(rows-1):
    if '.NS' in str(df.loc[t,'STOCK_NAME']) :
        df.loc[t,'REF_SYMBOL'] = df.loc[t,'STOCK_NAME']
# Save to Excel in the same directory as the script
output_file = Path(__file__).parent / "NSEBSE_stocks_name.xlsx"

try:
    df.to_excel(output_file, engine='openpyxl')
    print(f"File saved successfully at {output_file}")
except Exception as e:
    print(f"Error saving file: {e}")