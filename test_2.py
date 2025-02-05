import os
from binance.client import Client
import pandas as pd
import ta
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Retrieve the API keys from the environment variables
api_key = os.getenv('BINANCE_API_KEY')
api_secret = os.getenv('BINANCE_API_SECRET')
client = Client(api_key, api_secret)

def get_rsi(symbol, interval, limit=500):
    # Fetch klines (candlestick) data from Binance
    klines = client.get_klines(symbol=symbol, interval=interval, limit=limit)
    
    # Convert the data to a pandas DataFrame
    df = pd.DataFrame(klines, columns=['timestamp', 'open', 'high', 'low', 'close', 'volume', 
                                       'close_time', 'quote_asset_volume', 'number_of_trades', 
                                       'taker_buy_base_asset_volume', 'taker_buy_quote_asset_volume', 'ignore'])
    
    # Convert close prices to numeric values
    df['close'] = pd.to_numeric(df['close'])
    
    # Calculate RSI using the 'ta' library
    df['rsi'] = ta.momentum.RSIIndicator(df['close'], window=14).rsi()
    
    # Return the last 24 RSI values
    return df['rsi'].dropna().tail(24).values

with open('textcoin.txt', 'r') as file:
    symbols = [line.strip() for line in file.readlines()]

intervals = ['15m', '1h', '4h']  # Specify the intervals here

# Dictionary to store results for each interval
results = {interval: {'rsi_high': [], 'rsi_low': []} for interval in intervals}

for interval in intervals:
    for symbol in symbols:
        rsi_values = get_rsi(symbol, interval)
        for rsi_value in rsi_values:
            if rsi_value >= 80:
                results[interval]['rsi_high'].append({
                    'Tên': symbol,
                    'Chart URL': f'https://www.tradingview.com/chart/?symbol=BINANCE:{symbol}'
                })
                break
            elif rsi_value <= 20:
                results[interval]['rsi_low'].append({
                    'Tên': symbol,
                    'Chart URL': f'https://www.tradingview.com/chart/?symbol=BINANCE:{symbol}'
                })
                break

# Define the Excel filename
excel_file = 'rsi_filtered_data.xlsx'

# Check if the file exists
if os.path.exists(excel_file):
    # If the file exists, read the existing data
    with pd.ExcelWriter(excel_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        for interval in intervals:
            rsi_high_df = pd.DataFrame(results[interval]['rsi_high'])
            rsi_low_df = pd.DataFrame(results[interval]['rsi_low'])
            
            # Create a new sheet for the interval
            sheet_name = f'{interval}'
            start_row = 0
            
            # Write RSI >= 80 data
            rsi_high_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
            start_row += len(rsi_high_df) + 2  # Add 2 rows spacing
            
            # Write RSI <= 20 data
            rsi_low_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
    print(f'Filtered RSI data has been updated in {excel_file}')
else:
    # If the file does not exist, create it with the new data
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        for interval in intervals:
            rsi_high_df = pd.DataFrame(results[interval]['rsi_high'])
            rsi_low_df = pd.DataFrame(results[interval]['rsi_low'])
            
            # Create a new sheet for the interval
            sheet_name = f'{interval}'
            start_row = 0
            
            # Write RSI >= 80 data
            rsi_high_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
            start_row += len(rsi_high_df) + 2  # Add 2 rows spacing
            
            # Write RSI <= 20 data
            rsi_low_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
    print(f'Filtered RSI data has been exported to {excel_file}')
