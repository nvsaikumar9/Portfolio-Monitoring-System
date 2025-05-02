import pandas as pd

df = pd.read_excel('Portfolio_Analyser.xlsx')

#current gainer and looser
#the total portfolio value

for j in df['PORT_HOLDER'].unique():
    # Filter the DataFrame for the current portfolio holder
    df_port_ = df[df['PORT_HOLDER']== j]
    rows, columns = df_port_.shape

    Report = ''
    Report += f"Portfolio Holder: {j}\n"
    Report += f"Email : {df_port_['EMAIL'].iloc[0]}\n"

    # Calculate the total portfolio value
    total_portfolio_value = [(i*j) for i,j in zip(df_port_['AVG_PRICE'], df_port_['SHARES'])]
    total_portfolio_value = sum(total_portfolio_value)

    current_portfolio_value = [(i*j) for i,j in zip(df_port_['CLOSE'], df_port_['SHARES'])]
    current_portfolio_value = sum(current_portfolio_value)
    profitloss = current_portfolio_value - total_portfolio_value

    Top_looser_per = df_port_['ACTUAL_CLOSE%'].min()
    Top_looser_stock = df_port_['STOCK_NAME'][df_port_['ACTUAL_CLOSE%'].idxmin()]
    Top_gainer_stock = df_port_['STOCK_NAME'][df_port_['ACTUAL_CLOSE%'].idxmax()]
    Top_gainer_per = df_port_['ACTUAL_CLOSE%'].max()

    Report += f"Top Gainer: {df_port_['STOCK_NAME'][df_port_['ACTUAL_CLOSE%'].idxmax()]} with change of {Top_gainer_per:.2f}%\n"
    Report += f"Top Looser: {df_port_['STOCK_NAME'][df_port_['ACTUAL_CLOSE%'].idxmin()]} with change of {Top_looser_per:.2f}%\n"
    
    Report += f"Portfolio Holder: {j}, P/L : {profitloss:.2f}\n"
    Report += f"Total invested: {total_portfolio_value:.2f}\n"
    Report += f"Current Portfolio Value: {current_portfolio_value:.2f}\n"
    
    for i in range(rows):
        # Extract the stock name and its corresponding values
        stock_name = df_port_['STOCK_NAME'].iloc[i]
        avg_price = df_port_['AVG_PRICE'].iloc[i]
        shares = df_port_['SHARES'].iloc[i]
        threshold_limit = df_port_['THRESHOLD_LIMIT'].iloc[i]
        current_price = df_port_['CLOSE'].iloc[i]
        low_change = df_port_['%LOW_CHANGE'].iloc[i]
        high_change = df_port_['%HIGH_CHANGE'].iloc[i]
        current_price_change = df_port_['ACTUAL_CLOSE%'].iloc[i]

        # Calculate the weightage of the stock in the portfolio
        total_investment_share = avg_price * shares
        weightage = (total_investment_share / total_portfolio_value) * 100

        if threshold_limit < low_change:
            Report += f"Portfolio Holder: {j}, Stock: {stock_name}, crossed threshold and corrected by {low_change:.2f}%, Today's change : {current_price_change:.2f}%, Holding weightage : {weightage:.2f}%\n"
            #print(f"Portfolio Holder: {j}, Stock: {stock_name}, crossed threshold by {low_change:.2f}%, Current at : {current_price_change:.2f}%, Holding weightage : {weightage:.2f}%")
        elif threshold_limit < high_change:
            Report += f"Portfolio Holder: {j}, Stock: {stock_name}, crossed threshold and raise by {high_change:.2f}%, Today's change: {current_price_change:.2f}%, Holding weightage : {weightage:.2f}%\n"
            #print(f"Portfolio Holder: {j}, Stock: {stock_name}, crossed threshold by {high_change:.2f}%, Current at : {current_price_change:.2f}%, Holding weightage : {weightage:.2f}%")

    with open(f'Report.txt', 'w+') as f:
        f.write(Report)
        f.close()