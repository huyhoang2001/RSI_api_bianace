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

interval = '15m'  # Specify the interval here

# Lists to store symbols and their chart URLs that meet the RSI conditions
rsi_high = []  # RSI >= 80
rsi_low = []   # RSI <= 20

for symbol in symbols:
    rsi_values = get_rsi(symbol, interval)
    for rsi_value in rsi_values:
        if rsi_value >= 80:
            rsi_high.append({
                'Tên': symbol,
                'Chart URL': f'https://www.tradingview.com/chart/?symbol=BINANCE:{symbol}'
            })
            break
        elif rsi_value <= 20:
            rsi_low.append({
                'Tên': symbol,
                'Chart URL': f'https://www.tradingview.com/chart/?symbol=BINANCE:{symbol}'
            })
            break

# Convert the lists to DataFrames
rsi_high_df = pd.DataFrame(rsi_high)
rsi_low_df = pd.DataFrame(rsi_low)

# Define the Excel filename
excel_file = 'rsi_filtered_data.xlsx'

# Check if the file exists
if os.path.exists(excel_file):
    # If the file exists, read the existing data
    with pd.ExcelWriter(excel_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        # Write the new data to the respective sheets
        rsi_high_df.to_excel(writer, sheet_name='RSI >= 80', index=False)
        rsi_low_df.to_excel(writer, sheet_name='RSI <= 20', index=False)
    print(f'Filtered RSI data has been updated in {excel_file}')
else:
    # If the file does not exist, create it with the new data
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        rsi_high_df.to_excel(writer, sheet_name='RSI >= 80', index=False)
        rsi_low_df.to_excel(writer, sheet_name='RSI <= 20', index=False)
    print(f'Filtered RSI data has been exported to {excel_file}')
