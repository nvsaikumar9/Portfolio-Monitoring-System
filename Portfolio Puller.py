import openpyxl
import yfinance as yf
import pandas as pd
import os

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Construct relative paths for the required files
stocks_file = os.path.join(script_dir, 'NSEBSE_stocks_name.xlsx')
portfolio_file = os.path.join(script_dir, 'Portfolio_details.xlsx')

# Load the stock data from an Excel file
df = pd.read_excel(stocks_file)

# Initialize the portfolio dictionary
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
    """
    Collects portfolio details for a specific account holder.
    """
    while True:
        stock_Name = input('Enter the name of the holding stock (e.g., TCS, WIPRO). If done, type "Done": ').strip().upper()

        if stock_Name == 'DONE':
            break

        while True:
            bse_nse = input('Enter the stock exchange name (e.g., NSE, BSE): ').strip().upper()
            if bse_nse not in ['NSE', 'BSE']:
                print('Invalid input. Please enter either "NSE" or "BSE".')
                continue
            elif bse_nse == 'NSE':
                stock_Name_ = stock_Name + '.NS'
                break
            elif bse_nse == 'BSE':
                stock_Name_ = stock_Name + '.BO'
                break

        if stock_Name_ not in df['REF_SYMBOL'].values:
            print('Stock does not exist in the database. Please enter a valid stock name.')
            continue

        try:
            Avg_price = float(input('Enter the average price of the stock (e.g., 1500.50): '))
        except ValueError:
            print('Invalid input. Please enter a numeric value for the average price.')
            continue

        try:
            Lots = int(input('Enter the number of stocks held in the portfolio (e.g., 10): '))
        except ValueError:
            print('Invalid input. Please enter an integer value for the number of stocks.')
            continue

        # Append the entered details to the portfolio dictionary
        port_['PORT_HOLDER'].append(Port_Holder)
        port_['STOCK_NAME'].append(df[df['REF_SYMBOL'] == stock_Name_]['STOCK_NAME'].values[0])
        port_['AVG_PRICE'].append(Avg_price)
        port_['SHARES'].append(Lots)
        port_['THRESHOLD_LIMIT'].append(Thershold_limit)
        port_['Email'].append(Email)
        port_['BSE/NSE'].append(bse_nse)
        port_['REF_SYMBOL'].append(stock_Name_)

while True:
    Port_Holder = input('Enter the name of the account holder. If done, type "Done": ').strip().upper()
    if Port_Holder == 'DONE':
        break

    try:
        Thershold_limit = float(input('Enter the threshold limit for a stock (e.g., 5.0 for 5%): '))
    except ValueError:
        print('Invalid input. Please enter a numeric value for the threshold limit.')
        continue

    while True:
        Email = input('Enter the email of the portfolio holder: ').strip()
        if '@' not in Email or '.' not in Email.split('@')[1]:
            print('Invalid email address. Please enter a valid email address.')
            continue
        else:
            break

    Portfolio_details(Port_Holder, Thershold_limit, Email)

# Convert the portfolio dictionary to a DataFrame
df1 = pd.DataFrame(port_)

# Save the portfolio details to an Excel file using a relative path
df1.to_excel(portfolio_file, index=True, engine='openpyxl')

print("Portfolio details have been successfully saved to 'Portfolio_details.xlsx'.")