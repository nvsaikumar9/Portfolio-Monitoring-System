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

def send_email(sender_email, sender_password, recipient_email, subject, plain_body, html_body=None):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Add plain text and HTML content
    msg.attach(MIMEText(plain_body, 'plain'))
    if html_body:
        msg.attach(MIMEText(html_body, 'html'))

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        print(f"Email sent successfully to {recipient_email}!")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")

def main():
    try:
        nse = Nse()

        try:
            all_stock_symbols = nse.get_stock_codes()
            all_stock_symbols = [i+'.NS' for i in all_stock_symbols]
        except Exception as e:
            print(f"Error fetching stock codes: {e}")
            return

        data = {
            'STOCK_NAME': all_stock_symbols
        }

        try:
            df1 = pd.DataFrame(data)
            df1.to_excel(r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\NSE_stocks_name.xlsx', engine='openpyxl')
        except Exception as e:
            print(f"Error saving stock names to Excel: {e}")
            return

        # Read stock names from Excel
        stock_names = df1['STOCK_NAME'].tolist()

        # Initialize LTP dictionary
        LTP = {'STOCK_NAME': [], 'CLOSE': [], 'OPEN': [], 'HIGH': [], 'LOW': [], 'PREVIOUS_CLOSE': []}

        try:
            # Fetch data for all stocks in one API call
            price_data = yf.download(tickers=stock_names, period="2d", interval="1d", group_by='ticker', threads=True)
        except Exception as e:
            print(f"Error fetching stock data from Yahoo Finance: {e}")
            return

        # Process each stock
        for stock_name in stock_names:
            try:
                if stock_name in price_data:
                    stock_data = price_data[stock_name]

                    if not stock_data.empty:
                        # Extract the latest data
                        latest_data = stock_data.iloc[-1]
                        LTP['STOCK_NAME'].append(stock_name)
                        LTP['CLOSE'].append(latest_data['Close'])
                        LTP['OPEN'].append(latest_data['Open'])
                        LTP['HIGH'].append(latest_data['High'])
                        LTP['LOW'].append(latest_data['Low'])

                        last_data = stock_data.iloc[0]
                        LTP['PREVIOUS_CLOSE'].append(last_data['Close'])
            except Exception as e:
                print(f"Error processing stock {stock_name}: {e}")
                continue

        try:
            # Convert LTP dictionary to DataFrame and save to Excel
            df2 = pd.DataFrame(LTP)
            df2.to_excel(r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\API_stock_prices.xlsx', engine='openpyxl')
        except Exception as e:
            print(f"Error saving API stock prices to Excel: {e}")
            return

        try:
            # Process portfolio details and generate reports
            df_portfolio = pd.read_excel(r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\Portfolio_details.xlsx', engine='openpyxl')
            df_API = df2
        except Exception as e:
            print(f"Error reading portfolio or API data: {e}")
            return

        Uniq_port_holder = list(df_portfolio['PORT_HOLDER'].unique())

        required_columns_portfolio = ['PORT_HOLDER', 'STOCK_NAME', 'AVG_PRICE', 'SHARES']
        required_columns_api = ['STOCK_NAME', 'CLOSE', 'OPEN', 'HIGH', 'LOW']

        if not all(col in df_portfolio.columns for col in required_columns_portfolio):
            raise ValueError("Missing required columns in Portfolio_details.xlsx")

        if not all(col in df_API.columns for col in required_columns_api):
            raise ValueError("Missing required columns in API_stock_prices.xlsx")

        # Initialize variables
        New_columns = {
            'PORT_HOLDER': [], 'STOCK_NAME': [], 'AVG_PRICE': [], 'SHARES': [], 
            '%HIGH_CHANGE': [], '%LOW_CHANGE': [], 'CURRENT_%': [], 'HIGH_TO_LOW': [],
            'CLOSE': [], 'OPEN': [], 'HIGH': [], 'LOW': [], 'THRESHOLD_LIMIT': [], 'PREVIOUS_CLOSE': [], 'ACTUAL_CLOSE%': [], 'EMAIL': []
        }

        # Process data
        for i in Uniq_port_holder:
            try:
                df_port_holder = df_portfolio[df_portfolio['PORT_HOLDER'] == i]
                for j in df_port_holder['STOCK_NAME']:
                    try:
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
                        previous_close = float(df_API_stock_values['PREVIOUS_CLOSE'].iloc[0])

                        # Extract portfolio values
                        df_port_stock_vales = df_port_holder[df_port_holder['STOCK_NAME'] == j]
                        Average_price = float(df_port_stock_vales['AVG_PRICE'].iloc[0])
                        Num_shares = int(df_port_stock_vales['SHARES'].iloc[0])

                        # Calculate metrics
                        per_change_High = ((High_price - previous_close) / previous_close) * 100
                        per_change_Low = ((previous_close - Low_price) / previous_close) * 100
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
                        New_columns['PREVIOUS_CLOSE'].append(previous_close)
                        New_columns['ACTUAL_CLOSE%'].append(((close_price - previous_close) / previous_close) * 100)
                        New_columns['EMAIL'].append(df_port_stock_vales['Email'].iloc[0])
                    except Exception as e:
                        print(f"Error processing stock {j} for portfolio holder {i}: {e}")
                        continue
            except Exception as e:
                print(f"Error processing portfolio holder {i}: {e}")
                continue

        try:
            # Save to Excel
            df4 = pd.DataFrame(New_columns)
            df4.to_excel(r'C:\Vizual Studio Code\Python Programs\Project-PriceAlert\Portfolio_Analyser.xlsx', engine='openpyxl')
        except Exception as e:
            print(f"Error saving portfolio analysis to Excel: {e}")
            return

        # Generate and send reports
        for j in df4['PORT_HOLDER'].unique():
            try:
                df_port_ = df4[df4['PORT_HOLDER'] == j]
                rows, columns = df_port_.shape

                # Extract portfolio holder's email
                email = df_port_['EMAIL'].iloc[0]
                portfolio_holder_name = j

                # Calculate portfolio metrics
                total_portfolio_value = sum(df_port_['AVG_PRICE'] * df_port_['SHARES'])
                current_portfolio_value = sum(df_port_['CLOSE'] * df_port_['SHARES'])
                profitloss = current_portfolio_value - total_portfolio_value

                Top_looser_per = df_port_['ACTUAL_CLOSE%'].min()
                Top_looser_stock = df_port_['STOCK_NAME'][df_port_['ACTUAL_CLOSE%'].idxmin()]
                Top_gainer_stock = df_port_['STOCK_NAME'][df_port_['ACTUAL_CLOSE%'].idxmax()]
                Top_gainer_per = df_port_['ACTUAL_CLOSE%'].max()

                # Identify stocks that crossed the threshold
                threshold_stocks = df_port_[
                    (df_port_['CLOSE'] > df_port_['THRESHOLD_LIMIT']) |
                    (df_port_['CLOSE'] < df_port_['THRESHOLD_LIMIT'])
                ]

                # Plain text body
                plain_body = f"""
Hello {portfolio_holder_name},

Here is your daily portfolio update:

Summary:
- Total Portfolio Value: ₹{current_portfolio_value:.2f}
- Profit/Loss: ₹{profitloss:.2f}
- Top Gainer: {Top_gainer_stock} (+{Top_gainer_per:.2f}%)
- Top Loser: {Top_looser_stock} (-{Top_looser_per:.2f}%)
"""

                if not threshold_stocks.empty:
                    plain_body += "\nStocks that crossed the threshold:\n"
                    for idx, row in threshold_stocks.iterrows():
                        stock_name = row['STOCK_NAME']
                        close_price = row['CLOSE']
                        threshold_limit = row['THRESHOLD_LIMIT']
                        plain_body += f"- {stock_name}: Close Price ₹{close_price:.2f}, Threshold ₹{threshold_limit:.2f}\n"

                plain_body += "\nPlease check the attached detailed report for more information.\n\nBest regards,\nYour Portfolio Tracker"

                # HTML body with colors, emojis, and arrows
                html_body = f"""
<html>
<head>
    <style>
        body {{
            font-family: Arial, sans-serif;
            background-color: #f9f9f9;
            color: #333;
        }}
        h2 {{
            color: #4CAF50;
        }}
        p {{
            font-size: 16px;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin-top: 20px;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }}
        th {{
            background-color: #4CAF50;
            color: white;
        }}
        tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
        tr:hover {{
            background-color: #ddd;
        }}
        .profit {{
            color: #4CAF50;
            font-weight: bold;
        }}
        .loss {{
            color: #f44336;
            font-weight: bold;
        }}
    </style>
</head>
<body>
    <h2>📊 Daily Portfolio Report for {portfolio_holder_name}</h2>
    <p><strong>Total Portfolio Value:</strong> <span style="color: #4CAF50;">₹{current_portfolio_value:.2f}</span></p>
    <p><strong>Profit/Loss:</strong> 
        <span class="{ 'profit' if profitloss >= 0 else 'loss' }">
            { '📈' if profitloss >= 0 else '📉' } ₹{profitloss:.2f}
        </span>
    </p>
    <p><strong>Top Gainer:</strong> <span style="color: #4CAF50;">{Top_gainer_stock} (+{Top_gainer_per:.2f}%)</span> 🚀</p>
    <p><strong>Top Loser:</strong> <span style="color: #f44336;">{Top_looser_stock} (-{Top_looser_per:.2f}%)</span> 📉</p>
"""

                if not threshold_stocks.empty:
                    html_body += """
    <h3>⚠️ Stocks That Crossed the Threshold:</h3>
    <table>
        <tr>
            <th>Stock Name</th>
            <th>Close Price</th>
            <th>Threshold</th>
        </tr>
"""
                    for idx, row in threshold_stocks.iterrows():
                        stock_name = row['STOCK_NAME']
                        close_price = row['CLOSE']
                        threshold_limit = row['THRESHOLD_LIMIT']
                        html_body += f"""
        <tr>
            <td>{stock_name}</td>
            <td>₹{close_price:.2f}</td>
            <td>₹{threshold_limit:.2f}</td>
        </tr>
"""
                    html_body += """
    </table>
"""

                html_body += """
    <p>Thank you for using our service! 😊</p>
</body>
</html>
"""

                # Send email
                sender_email = 'n.v.saikumar9@gmail.com'
                sender_password = "upuuozxqebztquel"
                subject = f"Daily Portfolio Report for {portfolio_holder_name}"

                send_email(sender_email, sender_password, email, subject, plain_body, html_body)
                print(f"Email sent to {portfolio_holder_name} ({email})")
            except Exception as e:
                print(f"Error sending email to {portfolio_holder_name}: {e}")
                continue

    except Exception as e:
        print(f"An error occurred in the main function: {e}")

# Schedule the main function to run every 2 minutes
schedule.every(0.1).minutes.do(main)

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