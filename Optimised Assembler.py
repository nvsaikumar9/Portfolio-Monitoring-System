import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import schedule
import time
import pandas as pd
from nsetools import Nse
import openpyxl
import yfinance as yf
import os
import sys

def send_email(sender_email, sender_password, recipient_email, subject, body):
    """Send an email using SMTP."""
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        print(f"Email sent successfully to {recipient_email}!")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")

def fetch_stock_data(stock_names):
    """Fetch stock data using Yahoo Finance."""
    try:
        return yf.download(tickers=stock_names, period="2d", interval="1d", group_by='ticker', threads=True)
    except Exception as e:
        print(f"Error fetching stock data from Yahoo Finance: {e}")
        return None

def save_to_excel(dataframe, filepath):
    """Save a DataFrame to an Excel file."""
    try:
        dataframe.to_excel(filepath, engine='openpyxl')
    except Exception as e:
        print(f"Error saving data to Excel at {filepath}: {e}")

def process_portfolio_data(df_portfolio, df_API):
    """Process portfolio data and generate analysis."""
    required_columns_portfolio = ['PORT_HOLDER', 'STOCK_NAME', 'AVG_PRICE', 'SHARES']
    required_columns_api = ['STOCK_NAME', 'CLOSE', 'OPEN', 'HIGH', 'LOW']

    if not all(col in df_portfolio.columns for col in required_columns_portfolio):
        raise ValueError("Missing required columns in Portfolio_details.xlsx")

    if not all(col in df_API.columns for col in required_columns_api):
        raise ValueError("Missing required columns in API_stock_prices.xlsx")

    New_columns = {
        'PORT_HOLDER': [], 'STOCK_NAME': [], 'AVG_PRICE': [], 'SHARES': [],
        '%HIGH_CHANGE': [], '%LOW_CHANGE': [], 'CURRENT_%': [], 'HIGH_TO_LOW': [],
        'CLOSE': [], 'OPEN': [], 'HIGH': [], 'LOW': [], 'THRESHOLD_LIMIT': [], 'PREVIOUS_CLOSE': [], 'ACTUAL_CLOSE%': [], 'EMAIL': []
    }

    for port_holder in df_portfolio['PORT_HOLDER'].unique():
        df_port_holder = df_portfolio[df_portfolio['PORT_HOLDER'] == port_holder]
        for stock_name in df_port_holder['STOCK_NAME']:
            df_API_stock_values = df_API[df_API['STOCK_NAME'] == stock_name + '.NS']
            if df_API_stock_values.empty:
                print(f"Stock {stock_name+'.NS'} not found in API data. Skipping...")
                continue

            try:
                # Extract API values
                close_price = float(df_API_stock_values['CLOSE'].iloc[0])
                open_price = float(df_API_stock_values['OPEN'].iloc[0])
                high_price = float(df_API_stock_values['HIGH'].iloc[0])
                low_price = float(df_API_stock_values['LOW'].iloc[0])
                previous_close = float(df_API_stock_values['PREVIOUS_CLOSE'].iloc[0])

                # Extract portfolio values
                df_port_stock_values = df_port_holder[df_port_holder['STOCK_NAME'] == stock_name]
                avg_price = float(df_port_stock_values['AVG_PRICE'].iloc[0])
                num_shares = int(df_port_stock_values['SHARES'].iloc[0])
                threshold_limit = float(df_port_stock_values['THRESHOLD_LIMIT'].iloc[0])
                email = df_port_stock_values['Email'].iloc[0]

                # Calculate metrics
                per_change_high = ((high_price - previous_close) / previous_close) * 100
                per_change_low = ((previous_close - low_price) / previous_close) * 100
                per_current = ((close_price - open_price) / open_price) * 100
                delta = ((high_price - low_price) / low_price) * 100
                actual_close_percent = ((close_price - previous_close) / previous_close) * 100

                # Append to New_columns
                New_columns['PORT_HOLDER'].append(port_holder)
                New_columns['STOCK_NAME'].append(stock_name)
                New_columns['%HIGH_CHANGE'].append(per_change_high)
                New_columns['%LOW_CHANGE'].append(per_change_low)
                New_columns['CURRENT_%'].append(per_current)
                New_columns['HIGH_TO_LOW'].append(delta)
                New_columns['AVG_PRICE'].append(avg_price)
                New_columns['SHARES'].append(num_shares)
                New_columns['CLOSE'].append(close_price)
                New_columns['HIGH'].append(high_price)
                New_columns['LOW'].append(low_price)
                New_columns['OPEN'].append(open_price)
                New_columns['THRESHOLD_LIMIT'].append(threshold_limit)
                New_columns['PREVIOUS_CLOSE'].append(previous_close)
                New_columns['ACTUAL_CLOSE%'].append(actual_close_percent)
                New_columns['EMAIL'].append(email)
            except Exception as e:
                print(f"Error processing stock {stock_name} for portfolio holder {port_holder}: {e}")
                continue

    return pd.DataFrame(New_columns)

def generate_reports(df_analysis):
    """Generate reports for each portfolio holder."""
    reports = []
    for port_holder in df_analysis['PORT_HOLDER'].unique():
        try:
            df_port = df_analysis[df_analysis['PORT_HOLDER'] == port_holder]
            total_portfolio_value = sum(df_port['AVG_PRICE'] * df_port['SHARES'])
            current_portfolio_value = sum(df_port['CLOSE'] * df_port['SHARES'])
            profit_loss = current_portfolio_value - total_portfolio_value

            report = f"Portfolio Holder: {port_holder}\n"
            report += f"Email: {df_port['EMAIL'].iloc[0]}\n"
            report += f"Total Invested: {total_portfolio_value:.2f}\n"
            report += f"Current Portfolio Value: {current_portfolio_value:.2f}\n"
            report += f"Profit/Loss: {profit_loss:.2f}\n"

            for _, row in df_port.iterrows():
                stock_name = row['STOCK_NAME']
                weightage = (row['AVG_PRICE'] * row['SHARES'] / total_portfolio_value) * 100
                if row['THRESHOLD_LIMIT'] < row['%LOW_CHANGE']:
                    report += f"Stock {stock_name} corrected by {row['%LOW_CHANGE']:.2f}% (Threshold crossed).\n"
                elif row['THRESHOLD_LIMIT'] < row['%HIGH_CHANGE']:
                    report += f"Stock {stock_name} raised by {row['%HIGH_CHANGE']:.2f}% (Threshold crossed).\n"

            reports.append((df_port['EMAIL'].iloc[0], report))
        except Exception as e:
            print(f"Error generating report for portfolio holder {port_holder}: {e}")
            continue

    return reports

def main():
    try:
        nse = Nse()
        all_stock_symbols = [symbol + '.NS' for symbol in nse.get_stock_codes()]
        stock_data = fetch_stock_data(all_stock_symbols)
        if stock_data is None:
            return

        # Save stock data
        ltp_data = {
            'STOCK_NAME': stock_data.columns.levels[0],
            'CLOSE': stock_data.xs('Close', level=1, axis=1).iloc[-1].values,
            'OPEN': stock_data.xs('Open', level=1, axis=1).iloc[-1].values,
            'HIGH': stock_data.xs('High', level=1, axis=1).iloc[-1].values,
            'LOW': stock_data.xs('Low', level=1, axis=1).iloc[-1].values,
            'PREVIOUS_CLOSE': stock_data.xs('Close', level=1, axis=1).iloc[0].values
        }
        df_API = pd.DataFrame(ltp_data)
        save_to_excel(df_API, r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\API_stock_prices.xlsx')

        # Process portfolio
        df_portfolio = pd.read_excel(r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\Portfolio_details.xlsx', engine='openpyxl')
        df_analysis = process_portfolio_data(df_portfolio, df_API)
        save_to_excel(df_analysis, r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\Portfolio_Analyser.xlsx')

        # Generate and send reports
        reports = generate_reports(df_analysis)
        for email, report in reports:
            send_email('n.v.saikumar9@gmail.com', 'upuuozxqebztquel', email, "Daily Report on Portfolio", report)
    except Exception as e:
        print(f"An error occurred in the main function: {e}")

# Schedule the main function to run every 2 minutes
schedule.every(2).minutes.do(main)

print("Scheduler is running...")

# Keep the script running
try:
    while True:
        schedule.run_pending()
        time.sleep(1)
except KeyboardInterrupt:
    print("Scheduler stopped by user.")
except Exception as e:
    print(f"An error occurred in the scheduler: {e}")