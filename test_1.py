import os
import concurrent.futures
import pandas as pd
import numpy as np
import ta
import gspread
from binance.client import Client
from dotenv import load_dotenv
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from colorama import init
from oauth2client.service_account import ServiceAccountCredentials

init(autoreset=True)
load_dotenv()

# Danh sách khung thời gian cho phép
ALLOWED_INTERVALS = ["15m", "30m", "1h", "4h", "1d"]

def ask_rsi_period():
    while True:
        user_input = input("Nhập RSI period (số nguyên, ví dụ 14; nhập 00 để quay lại hoặc thoát): ").strip()
        if user_input == "00":
            print("Thoát tùy chỉnh RSI period.")
            return None
        try:
            period = int(user_input)
            if period <= 0:
                raise ValueError
            return period
        except ValueError:
            print("Vui lòng nhập một số nguyên dương hợp lệ.")

def ask_intervals():
    while True:
        print("\nNhập khung thời gian muốn tính RSI. Cho phép nhập một hoặc nhiều giá trị cách nhau bằng dấu cách.")
        print("Các lựa chọn có thể: " + " ".join(ALLOWED_INTERVALS))
        print("Nhập 00 để quay lại bước trước.")
        user_input = input("Nhập khung thời gian của bạn: ").strip()
        if user_input == "00":
            return None
        # Tách các giá trị
        intervals = user_input.split()
        # Kiểm tra từng giá trị có trong danh sách cho phép
        valid = True
        for i in intervals:
            if i not in ALLOWED_INTERVALS:
                print(f"Khung thời gian '{i}' không hợp lệ.")
                valid = False
                break
        if valid and len(intervals) > 0:
            return intervals
        else:
            print("Vui lòng nhập lại các khung thời gian hợp lệ.")

class BinanceRSIAnalyzer:
    def __init__(self):
        self.api_key = os.getenv('BINANCE_API_KEY')
        self.api_secret = os.getenv('BINANCE_API_SECRET')
        self.client = Client(self.api_key, self.api_secret)
        self.symbols = self._load_symbols()
        self.intervals = ['15m','1h', '4h','1d']  # giá trị mặc định, sẽ được thay đổi theo input
        self.rsi_period = 14  # mặc định 14, sẽ cập nhật theo input
        self.excel_file = 'rsi_filtered_data.xlsx'

    def _load_symbols(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(current_dir, 'textcoin.txt')
        with open(file_path, 'r') as file:
            return [line.strip() for line in file.readlines()]

    def _calculate_rsi(self, close_prices, window):
        rsi = ta.momentum.RSIIndicator(close_prices, window=window).rsi()
        return rsi.dropna().tail(5)

    def _fetch_and_process_data(self, symbol, interval):
        try:
            klines = self.client.get_klines(
                symbol=symbol,
                interval=interval,
                limit=500
            )
            close_prices = pd.Series([float(k[4]) for k in klines])
            if len(close_prices) < self.rsi_period:
                return None

            rsi = self._calculate_rsi(close_prices, self.rsi_period)
            return (symbol, interval, rsi)
        except Exception as e:
            print(f"Error processing {symbol} {interval}: {str(e)}")
            return None

    def _process_result(self, results):
        output = {interval: {'rsi_high': [], 'rsi_low': []} for interval in self.intervals}
        for symbol, interval, rsi in results:
            if (rsi >= 80).any():
                output[interval]['rsi_high'].append({
                    'Tên': symbol,
                    'Chart URL': f'https://www.tradingview.com/chart/?symbol=BINANCE:{symbol}'
                })
            elif (rsi <= 20).any():
                output[interval]['rsi_low'].append({
                    'Tên': symbol,
                    'Chart URL': f'https://www.tradingview.com/chart/?symbol=BINANCE:{symbol}'
                })
        return output

    def _save_to_excel(self, data):   
        with pd.ExcelWriter(self.excel_file, engine='openpyxl') as writer:
            for interval in self.intervals:
                high_df = pd.DataFrame(data[interval]['rsi_high'])
                low_df = pd.DataFrame(data[interval]['rsi_low'])
                spacer = pd.DataFrame([[''] * 3], columns=['Tên', 'Chart URL', 'RSI Condition'])
                combined_df = pd.concat([
                    high_df.assign(**{'RSI Condition': '>=80'}),
                    spacer,
                    low_df.assign(**{'RSI Condition': '<=20'})
                ], ignore_index=True)
                combined_df.to_excel(writer, sheet_name=interval, index=False, startrow=0)
                ws = writer.sheets[interval]
                for col in ws.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2) * 1.2
                    ws.column_dimensions[column].width = adjusted_width
                border_style = Border(left=Side(style='thin'),
                                        right=Side(style='thin'),
                                        top=Side(style='thin'),
                                        bottom=Side(style='thin'))
                for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                    for cell in row:
                        cell.border = border_style
                header_font = Font(bold=True, color='FFFFFF')
                header_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
                for cell in ws[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal='center')
                ws.auto_filter.ref = ws.dimensions
                ws.freeze_panes = 'A2'
        print(f'File Excel đã được cập nhật với định dạng chuyên nghiệp!')

    def _upload_to_google_sheet(self, data):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        creds_file = os.path.join(current_dir, '(name.json for user)')
        if not os.path.exists(creds_file):
            print(f"🚫 File JSON credentials không tồn tại: {creds_file}")
            return
        
        scope = ["https://spreadsheets.google.com/feeds",
                 "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
        client_sheet = gspread.authorize(creds)
        spreadsheet = client_sheet.open("RSI Data")
        for interval, items in data.items():
            try:
                worksheet = spreadsheet.worksheet(interval)
            except gspread.exceptions.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(title=interval, rows="100", cols="20")
            headers = ["Tên", "Chart URL", "RSI Condition"]
            values = [headers]
            for row in items.get('rsi_high', []):
                values.append([row.get("Tên"), row.get("Chart URL"), ">=80"])
            values.append([""] * len(headers))
            for row in items.get('rsi_low', []):
                values.append([row.get("Tên"), row.get("Chart URL"), "<=20"])
            worksheet.clear()
            worksheet.update("A1", values)
        print("✅ Dữ liệu đã được upload lên Google Sheet thành công!")

    def analyze(self):
        # Hỏi người dùng RSI period
        while True:
            period = ask_rsi_period()
            if period is not None:
                self.rsi_period = period
                break
            else:
                print("Cần nhập RSI period để tiếp tục.")
                
        # Hỏi người dùng khung thời gian (có thể nhập đơn lẻ hoặc mảng)
        while True:
            intervals = ask_intervals()
            if intervals is not None:
                self.intervals = intervals
                break
            else:
                print("Cần nhập khung thời gian hợp lệ để tiếp tục.")

        print(f"\nĐang tính RSI sử dụng period = {self.rsi_period}")
        print(f"Khung thời gian được chọn: {self.intervals}\n")

        results = []
        for interval in self.intervals:
            print(f'🔄 Đang xử lý khung thời gian {interval}...')
            with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
                futures = [executor.submit(self._fetch_and_process_data, symbol, interval) 
                           for symbol in self.symbols]
                completed = 0
                total = len(self.symbols)
                while futures:
                    done, futures = concurrent.futures.wait(futures, return_when=concurrent.futures.FIRST_COMPLETED)
                    completed += len(done)
                    progress = (completed / total) * 100
                    print(f'\r📊 Tiến trình {interval}: {progress:.1f}%', end='', flush=True)
                    for future in done:
                        result = future.result()
                        if result:
                            results.append(result)
            print(f'\n✅ Đã hoàn thành khung {interval}!')

        processed_data = self._process_result(results)
        self._save_to_excel(processed_data)
        self._upload_to_google_sheet(processed_data)
        print(f'\n🔥 Tất cả dữ liệu đã được lưu vào {self.excel_file}')

if __name__ == "__main__":
    analyzer = BinanceRSIAnalyzer()
    analyzer.analyze()
