import os
from binance.client import Client
import pandas as pd
import ta
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side

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
    
    # Return the last 5 RSI values
    return df['rsi'].dropna().tail(5).values

with open('textcoin.txt', 'r') as file:
    symbols = [line.strip() for line in file.readlines()]

interval = '1h'  # Specify the interval here

# Create a DataFrame to store RSI values and chart URLs
data = {'Tên': symbols, 'Chart URL': []}
for i in range(1, 6):
    data[f'rsi{i}'] = []

for symbol in symbols:
    rsi_values = get_rsi(symbol, interval)
    data_row = {'Tên': symbol, 'Chart URL': f'https://www.tradingview.com/chart/?symbol=BINANCE:{symbol}'}
    for i, rsi_value in enumerate(rsi_values, start=1):
        data_row[f'rsi{i}'] = rsi_value
    data['Chart URL'].append(data_row['Chart URL'])
    for i in range(1, 6):
        data[f'rsi{i}'].append(data_row[f'rsi{i}'])

# Convert data to DataFrame
rsi_df = pd.DataFrame(data)

# Define the Excel filename
excel_file = 'rsi_data.xlsx'

# Check if the file exists
if os.path.exists(excel_file):
    # If the file exists, read the existing data
    existing_df = pd.read_excel(excel_file, index_col=0)
    
    # Concatenate the new data with the existing data
    updated_df = pd.concat([existing_df, rsi_df], ignore_index=True)
    
    # Save the updated DataFrame back to the Excel file
    updated_df.to_excel(excel_file, index=False)
    print(f'RSI data has been updated in {excel_file}')
else:
    # If the file does not exist, create it with the new data
    rsi_df.to_excel(excel_file, index=False)
    print(f'RSI data has been exported to {excel_file}')

# Apply conditional formatting
wb = load_workbook(excel_file)
ws = wb.active

# Define the fill colors
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
grays_fill = PatternFill(start_color="B6B7AF", end_color="FFA500", fill_type="solid")

# Convert RSI values to numeric values in the Excel sheet
for row in ws.iter_rows(min_row=2, min_col=3, max_col=7):  # Update columns range
    for cell in row:
        try:
            cell.value = float(cell.value)
        except (TypeError, ValueError):
            cell.value = None

# Apply the conditional formatting to the RSI columns
for row in ws.iter_rows(min_row=2, min_col=3, max_col=7):  # Update columns range
    for cell in row:
        if cell.value is not None:
            if cell.value >= 70:
                cell.fill = orange_fill
            elif cell.value <= 30:
                cell.fill = yellow_fill

# Add headers for the new columns
ws['H1'] = 'RSI >= 70'
ws['I1'] = 'RSI <= 30'

# Add conditional formulas to the sheet
for idx, row in enumerate(ws.iter_rows(min_row=2, min_col=1, max_col=8, values_only=True), start=2):
    ws[f'H{idx}'] = f'=IF(OR(C{idx}>=70,D{idx}>=70,E{idx}>=70,F{idx}>=70,G{idx}>=70),A{idx},"")'
    ws[f'I{idx}'] = f'=IF(OR(C{idx}<=30,D{idx}<=30,E{idx}<=30,F{idx}<=30,G{idx}<=30),A{idx},"")'

# Define the border style
border_style = Border(left=Side(style='thin'), 
                      right=Side(style='thin'), 
                      top=Side(style='thin'), 
                      bottom=Side(style='thin'))

# Apply the border to all cells
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    for cell in row:
        cell.border = border_style      

# Save the workbook
wb.save(excel_file)
print(f'Conditional formatting applied to {excel_file}')
