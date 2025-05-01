import openpyxl
import pandas as pd
import datetime

# Load Excel files
df_portfolio = pd.read_excel('Portfolio_details.xlsx')
df_API = pd.read_excel('API_stock_prices.xlsx')

# Validate required columns
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
df1 = pd.DataFrame(New_columns)
output_file = f'Portfolio_Analyser.xlsx'
df1.to_excel(output_file, index=False, engine='openpyxl')
