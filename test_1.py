import os
import concurrent.futures
import pandas as pd
import numpy as np
import ta
import gspread
import requests
from binance.client import Client
from dotenv import load_dotenv
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
from oauth2client.service_account import ServiceAccountCredentials

load_dotenv()

ALLOWED_INTERVALS = ["5m", "15m", "30m", "1h", "4h", "1d", "1w"]


def ask_rsi_period():
    while True:
        user_input = input("Nhập RSI period (nhập 00 để thoát): ").strip()
        if user_input == "00":
            return None
        try:
            period = int(user_input)
            if period <= 0:
                raise ValueError
            return period
        except ValueError:
            print("Vui lòng nhập số nguyên dương.")


def ask_intervals():
    while True:
        print(f"\nKhung thời gian cho phép: {' '.join(ALLOWED_INTERVALS)}")
        user_input = input("Nhập khung thời gian (00 để quay lại): ").strip()
        if user_input == "00":
            return None
        intervals = user_input.split()
        if all(i in ALLOWED_INTERVALS for i in intervals) and intervals:
            return intervals
        print("Khung thời gian không hợp lệ.")


def ask_analysis_mode():
    while True:
        print("\n" + "="*50)
        print("1. RSI cơ bản (RSI >= 80 hoặc <= 20)")
        print("2. RSI Divergence (Phân kỳ theo rules mới)")
        print("3. Cả hai")
        print("00. Thoát")
        print("="*50)
        choice = input("Chọn chế độ: ").strip()
        if choice == "00":
            return None
        if choice in ["1", "2", "3"]:
            return int(choice)
        print("Lựa chọn không hợp lệ.")


class BinanceRSIAnalyzer:
    def __init__(self):
        self.api_key = os.getenv('BINANCE_API_KEY')
        self.api_secret = os.getenv('BINANCE_API_SECRET')
        self.telegram_token = os.getenv('TELEGRAM_BOT_TOKEN')
        self.telegram_chat_id = os.getenv('TELEGRAM_CHAT_ID')
        self.client = Client(self.api_key, self.api_secret)
        self.symbols = self._load_symbols()
        self.intervals = ['15m', '1h', '4h', '1d']
        self.rsi_period = 14
        self.excel_file = 'rsi_filtered_data.xlsx'
        self.analysis_mode = 1
        
        self.oversold = 20
        self.overbought = 80
        self.middle_low = 40
        self.middle_high = 60
        
        self.min_candle_distance = 24
        self.max_candle_distance = 34
        self.scan_candles = 100

    def _load_symbols(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(current_dir, 'textcoin.txt')
        with open(file_path, 'r') as file:
            return [line.strip() for line in file.readlines()]

    def _calculate_rsi(self, close_prices, window):
        rsi = ta.momentum.RSIIndicator(close_prices, window=window).rsi()
        return rsi.dropna()

    def _find_local_extremes(self, rsi_values, price_values, window=3):
        rsi = np.array(rsi_values)
        prices = np.array(price_values)
        highs = []
        lows = []
        
        for i in range(window, len(rsi) - window):
            if rsi[i] == min(rsi[i-window:i+window+1]):
                lows.append({'index': i, 'rsi': rsi[i], 'price': prices[i]})
            if rsi[i] == max(rsi[i-window:i+window+1]):
                highs.append({'index': i, 'rsi': rsi[i], 'price': prices[i]})
        
        return highs, lows

    def _detect_divergence(self, rsi_series, price_series):
        if len(rsi_series) < self.scan_candles:
            return []
        
        rsi = np.array(rsi_series[-self.scan_candles:])
        prices = np.array(price_series[-self.scan_candles:])
        
        highs, lows = self._find_local_extremes(rsi, prices, window=3)
        results = []
        
        bearish = self._detect_bearish_divergence(rsi, prices, highs, lows)
        if bearish:
            results.append(bearish)
        
        bullish = self._detect_bullish_divergence(rsi, prices, highs, lows)
        if bullish:
            results.append(bullish)
        
        return results

    def _detect_bearish_divergence(self, rsi, prices, highs, lows):
        overbought_highs = [h for h in highs if h['rsi'] > self.overbought]
        
        if not overbought_highs:
            return None
        
        for high1 in reversed(overbought_highs):
            retracement_found = False
            retracement_index = None
            
            for i in range(high1['index'] + 1, len(rsi)):
                if self.middle_low < rsi[i] < self.middle_high:
                    retracement_found = True
                    retracement_index = i
                    break
            
            if not retracement_found:
                continue
            
            for high2 in highs:
                if high2['index'] <= retracement_index:
                    continue
                
                if not (self.middle_high < high2['rsi'] < self.overbought):
                    continue
                
                candle_distance = high2['index'] - high1['index']
                if candle_distance < self.min_candle_distance or candle_distance > self.max_candle_distance:
                    continue
                
                if high2['price'] > high1['price']:
                    current_index = len(rsi) - 1
                    distance_from_high2 = current_index - high2['index']
                    
                    if distance_from_high2 <= 5:
                        stage = 'CONFIRMED'
                    elif distance_from_high2 <= 15:
                        stage = 'DEVELOPING'
                    else:
                        stage = 'FORMING'
                    
                    return {
                        'type': 'bearish',
                        'stage': stage,
                        'strength': self._calculate_strength(stage, True)
                    }
        
        return None

    def _detect_bullish_divergence(self, rsi, prices, highs, lows):
        oversold_lows = [l for l in lows if l['rsi'] < self.oversold]
        
        if not oversold_lows:
            return None
        
        for low1 in reversed(oversold_lows):
            retracement_found = False
            retracement_index = None
            
            for i in range(low1['index'] + 1, len(rsi)):
                if self.middle_low < rsi[i] < self.middle_high:
                    retracement_found = True
                    retracement_index = i
                    break
            
            if not retracement_found:
                continue
            
            for low2 in lows:
                if low2['index'] <= retracement_index:
                    continue
                
                if not (self.oversold < low2['rsi'] < self.middle_low):
                    continue
                
                candle_distance = low2['index'] - low1['index']
                if candle_distance < self.min_candle_distance or candle_distance > self.max_candle_distance:
                    continue
                
                if low2['price'] < low1['price']:
                    current_index = len(rsi) - 1
                    distance_from_low2 = current_index - low2['index']
                    
                    if distance_from_low2 <= 5:
                        stage = 'CONFIRMED'
                    elif distance_from_low2 <= 15:
                        stage = 'DEVELOPING'
                    else:
                        stage = 'FORMING'
                    
                    return {
                        'type': 'bullish',
                        'stage': stage,
                        'strength': self._calculate_strength(stage, True)
                    }
        
        return None

    def _calculate_strength(self, stage, has_divergence):
        base_strength = {'FORMING': 25, 'DEVELOPING': 55, 'CONFIRMED': 85}
        strength = base_strength.get(stage, 0)
        if has_divergence:
            strength += 15
        return min(strength, 100)

    def _fetch_and_process_data(self, symbol, interval):
        try:
            klines = self.client.get_klines(symbol=symbol, interval=interval, limit=200)
            close_prices = pd.Series([float(k[4]) for k in klines])
            current_price = float(klines[-1][4])
            
            if len(close_prices) < self.rsi_period + self.scan_candles:
                return None
            
            rsi = self._calculate_rsi(close_prices, self.rsi_period)
            aligned_prices = close_prices.iloc[-len(rsi):].reset_index(drop=True)
            rsi_last5 = rsi.tail(5)
            
            divergences = self._detect_divergence(rsi, aligned_prices)
            
            bullish_div = None
            bearish_div = None
            for div in divergences:
                if div['type'] == 'bullish':
                    bullish_div = div
                elif div['type'] == 'bearish':
                    bearish_div = div
            
            return {
                'symbol': symbol,
                'interval': interval,
                'rsi_last5': rsi_last5,
                'current_rsi': round(rsi.iloc[-1], 2),
                'current_price': current_price,
                'divergence_bullish': bullish_div,
                'divergence_bearish': bearish_div
            }
        except Exception:
            return None

    def _process_result(self, results):
        output = {interval: {
            'rsi_high': [], 'rsi_low': [],
            'div_bullish_confirmed': [], 'div_bullish_developing': [], 'div_bullish_forming': [],
            'div_bearish_confirmed': [], 'div_bearish_developing': [], 'div_bearish_forming': []
        } for interval in self.intervals}
        
        for result in results:
            symbol = result['symbol']
            interval = result['interval']
            rsi_last5 = result['rsi_last5']
            chart_url = f'https://www.tradingview.com/chart/?symbol=BINANCE:{symbol}'
            
            if self.analysis_mode in [1, 3]:
                if (rsi_last5 >= 80).any():
                    output[interval]['rsi_high'].append({
                        'Tên': symbol, 
                        'Loại': 'RSI ≥ 80', 
                        'Chart URL': chart_url
                    })
                elif (rsi_last5 <= 20).any():
                    output[interval]['rsi_low'].append({
                        'Tên': symbol, 
                        'Loại': 'RSI ≤ 20', 
                        'Chart URL': chart_url
                    })
            
            if self.analysis_mode in [2, 3]:
                if result['divergence_bullish']:
                    div = result['divergence_bullish']
                    key = f'div_bullish_{div["stage"].lower()}'
                    output[interval][key].append({
                        'Tên': symbol,
                        'Loại': 'Bullish Div',
                        'Giai đoạn': div['stage'],
                        'Độ mạnh': f"{div['strength']}%",
                        'Chart URL': chart_url
                    })
                
                if result['divergence_bearish']:
                    div = result['divergence_bearish']
                    key = f'div_bearish_{div["stage"].lower()}'
                    output[interval][key].append({
                        'Tên': symbol,
                        'Loại': 'Bearish Div',
                        'Giai đoạn': div['stage'],
                        'Độ mạnh': f"{div['strength']}%",
                        'Chart URL': chart_url
                    })
        
        return output

    def _send_telegram_message(self, data):
        print(f"\n🔍 Debug Telegram:")
        print(f"   Token: {'✓ Có' if self.telegram_token else '✗ Không có'}")
        print(f"   Chat ID: {'✓ Có' if self.telegram_chat_id else '✗ Không có'}")
        
        if not self.telegram_token or not self.telegram_chat_id:
            print("   ⚠️ Thiếu TELEGRAM_BOT_TOKEN hoặc TELEGRAM_CHAT_ID trong .env")
            return False
        
        message_parts = [f"🔔 <b>RSI Alert - Period {self.rsi_period}</b>"]
        mode_text = {1: "RSI Cơ bản", 2: "RSI Divergence", 3: "RSI + Divergence"}
        message_parts.append(f"📊 Mode: {mode_text.get(self.analysis_mode)}\n")
        
        stage_emoji = {'CONFIRMED': '✅', 'DEVELOPING': '🔄', 'FORMING': '🌱'}
        has_any_data = False
        
        for interval in self.intervals:
            interval_data = data[interval]
            
            interval_has_data = any([
                interval_data['rsi_high'], 
                interval_data['rsi_low'],
                interval_data['div_bullish_confirmed'], 
                interval_data['div_bullish_developing'],
                interval_data['div_bullish_forming'], 
                interval_data['div_bearish_confirmed'],
                interval_data['div_bearish_developing'], 
                interval_data['div_bearish_forming']
            ])
            
            if interval_has_data:
                has_any_data = True
                message_parts.append(f"\n⏰ <b>Khung {interval}</b>")
                
                if self.analysis_mode in [1, 3]:
                    if interval_data['rsi_high']:
                        message_parts.append(f"\n📈 <b>RSI ≥ 80:</b>")
                        for item in interval_data['rsi_high'][:10]:
                            message_parts.append(f"• {item['Tên']} | <a href='{item['Chart URL']}'>Chart</a>")
                    
                    if interval_data['rsi_low']:
                        message_parts.append(f"\n📉 <b>RSI ≤ 20:</b>")
                        for item in interval_data['rsi_low'][:10]:
                            message_parts.append(f"• {item['Tên']} | <a href='{item['Chart URL']}'>Chart</a>")
                
                if self.analysis_mode in [2, 3]:
                    for stage in ['confirmed', 'developing', 'forming']:
                        key = f'div_bullish_{stage}'
                        if interval_data[key]:
                            message_parts.append(f"\n🟢 <b>BULLISH DIV - {stage_emoji[stage.upper()]} {stage.upper()}:</b>")
                            for item in interval_data[key][:10]:
                                message_parts.append(f"• {item['Tên']} | {item['Độ mạnh']} | <a href='{item['Chart URL']}'>Chart</a>")
                    
                    for stage in ['confirmed', 'developing', 'forming']:
                        key = f'div_bearish_{stage}'
                        if interval_data[key]:
                            message_parts.append(f"\n🔴 <b>BEARISH DIV - {stage_emoji[stage.upper()]} {stage.upper()}:</b>")
                            for item in interval_data[key][:10]:
                                message_parts.append(f"• {item['Tên']} | {item['Độ mạnh']} | <a href='{item['Chart URL']}'>Chart</a>")
        
        if not has_any_data:
            message_parts.append("\n✅ Không có tín hiệu nào")
        
        message = "\n".join(message_parts)
        
        if len(message) > 4000:
            message = message[:4000] + "\n\n... (truncated)"
        
        url = f"https://api.telegram.org/bot{self.telegram_token}/sendMessage"
        payload = {
            "chat_id": self.telegram_chat_id,
            "text": message,
            "parse_mode": "HTML",
            "disable_web_page_preview": True
        }
        
        try:
            print(f"   📤 Đang gửi message ({len(message)} ký tự)...")
            response = requests.post(url, json=payload, timeout=30)
            
            if response.status_code == 200:
                print("   ✅ Telegram: Gửi thành công!")
                return True
            else:
                print(f"   ❌ Telegram Error: Status {response.status_code}")
                print(f"   Response: {response.text}")
                return False
                
        except requests.exceptions.Timeout:
            print("   ❌ Telegram: Timeout")
            return False
        except requests.exceptions.ConnectionError:
            print("   ❌ Telegram: Connection Error")
            return False
        except Exception as e:
            print(f"   ❌ Telegram Exception: {str(e)}")
            return False

    def _save_to_excel(self, data):
        with pd.ExcelWriter(self.excel_file, engine='openpyxl') as writer:
            for interval in self.intervals:
                rows = []
                
                if self.analysis_mode in [1, 3]:
                    for item in data[interval]['rsi_high']:
                        rows.append({
                            'Tên': item['Tên'], 
                            'Loại': item['Loại'], 
                            'Giai đoạn': '-', 
                            'Độ mạnh': '-', 
                            'Chart URL': item['Chart URL']
                        })
                    for item in data[interval]['rsi_low']:
                        rows.append({
                            'Tên': item['Tên'], 
                            'Loại': item['Loại'], 
                            'Giai đoạn': '-', 
                            'Độ mạnh': '-', 
                            'Chart URL': item['Chart URL']
                        })
                
                if self.analysis_mode in [2, 3]:
                    for stage in ['confirmed', 'developing', 'forming']:
                        for item in data[interval][f'div_bullish_{stage}']:
                            rows.append({
                                'Tên': item['Tên'], 
                                'Loại': item['Loại'], 
                                'Giai đoạn': item['Giai đoạn'],
                                'Độ mạnh': item['Độ mạnh'], 
                                'Chart URL': item['Chart URL']
                            })
                        for item in data[interval][f'div_bearish_{stage}']:
                            rows.append({
                                'Tên': item['Tên'], 
                                'Loại': item['Loại'], 
                                'Giai đoạn': item['Giai đoạn'],
                                'Độ mạnh': item['Độ mạnh'], 
                                'Chart URL': item['Chart URL']
                            })
                
                columns = ['Tên', 'Loại', 'Giai đoạn', 'Độ mạnh', 'Chart URL']
                df = pd.DataFrame(rows, columns=columns) if rows else pd.DataFrame(columns=columns)
                df.to_excel(writer, sheet_name=interval, index=False)
                
                ws = writer.sheets[interval]
                for col in ws.columns:
                    max_length = max(len(str(cell.value or '')) for cell in col)
                    ws.column_dimensions[col[0].column_letter].width = min((max_length + 2) * 1.2, 80)
                
                border_style = Border(
                    left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='thin')
                )
                for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                    for cell in row:
                        cell.border = border_style
                
                header_font = Font(bold=True, color='FFFFFF')
                header_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
                for cell in ws[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal='center')
                
                stage_fills = {
                    'CONFIRMED': PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid'),
                    'DEVELOPING': PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid'),
                    'FORMING': PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
                }
                
                header_row = [cell.value for cell in ws[1]]
                stage_col = header_row.index('Giai đoạn') + 1 if 'Giai đoạn' in header_row else None
                
                for row_idx in range(2, ws.max_row + 1):
                    if stage_col:
                        stage_cell = ws.cell(row=row_idx, column=stage_col)
                        if stage_cell.value in stage_fills:
                            for col_idx in range(1, ws.max_column + 1):
                                ws.cell(row=row_idx, column=col_idx).fill = stage_fills[stage_cell.value]
                
                ws.auto_filter.ref = ws.dimensions
                ws.freeze_panes = 'A2'
        
        print(f'✅ Excel saved: {self.excel_file}')

    def _upload_to_google_sheet(self, data):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        creds_file = os.path.join(current_dir, 'your.json')
        
        if not os.path.exists(creds_file):
            print("⚠️ Google Sheet credentials not found")
            return
        
        try:
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
            client_sheet = gspread.authorize(creds)
            spreadsheet = client_sheet.open("RSI Data")
            
            for interval in self.intervals:
                try:
                    worksheet = spreadsheet.worksheet(interval)
                except gspread.exceptions.WorksheetNotFound:
                    worksheet = spreadsheet.add_worksheet(title=interval, rows="100", cols="20")
                
                headers = ["Tên", "Loại", "Giai đoạn", "Độ mạnh", "Chart URL"]
                values = [headers]
                
                if self.analysis_mode in [1, 3]:
                    for item in data[interval].get('rsi_high', []):
                        values.append([item['Tên'], item['Loại'], '-', '-', item['Chart URL']])
                    for item in data[interval].get('rsi_low', []):
                        values.append([item['Tên'], item['Loại'], '-', '-', item['Chart URL']])
                
                if self.analysis_mode in [2, 3]:
                    for stage in ['confirmed', 'developing', 'forming']:
                        for item in data[interval].get(f'div_bullish_{stage}', []):
                            values.append([
                                item['Tên'], item['Loại'], item['Giai đoạn'],
                                item['Độ mạnh'], item['Chart URL']
                            ])
                        for item in data[interval].get(f'div_bearish_{stage}', []):
                            values.append([
                                item['Tên'], item['Loại'], item['Giai đoạn'],
                                item['Độ mạnh'], item['Chart URL']
                            ])
                
                worksheet.clear()
                worksheet.update("A1", values)
            
            print("✅ Google Sheet updated!")
        except Exception as e:
            print(f"❌ Google Sheet error: {str(e)}")

    def analyze(self):
        mode = ask_analysis_mode()
        if mode is None:
            return
        self.analysis_mode = mode
        
        while True:
            period = ask_rsi_period()
            if period is not None:
                self.rsi_period = period
                break
        
        while True:
            intervals = ask_intervals()
            if intervals is not None:
                self.intervals = intervals
                break
        
        print(f"\n{'='*60}")
        print(f"📊 Mode: {self.analysis_mode} | RSI: {self.rsi_period} | Intervals: {self.intervals}")
        print(f"{'='*60}\n")
        
        results = []
        for interval in self.intervals:
            print(f'🔄 Processing {interval}...')
            with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
                futures = {
                    executor.submit(self._fetch_and_process_data, symbol, interval): symbol 
                    for symbol in self.symbols
                }
                completed = 0
                total = len(self.symbols)
                for future in concurrent.futures.as_completed(futures):
                    completed += 1
                    print(f'\r📊 {interval}: {(completed/total)*100:.1f}%', end='', flush=True)
                    result = future.result()
                    if result:
                        results.append(result)
            print(f'\n✅ Done {interval}!')
        
        processed_data = self._process_result(results)
        
        print(f"\n{'='*70}")
        print("📊 SUMMARY:")
        for interval in self.intervals:
            print(f"\n⏰ {interval}:")
            if self.analysis_mode in [1, 3]:
                print(f"   📈 RSI ≥ 80: {len(processed_data[interval]['rsi_high'])}")
                print(f"   📉 RSI ≤ 20: {len(processed_data[interval]['rsi_low'])}")
            if self.analysis_mode in [2, 3]:
                bull_c = len(processed_data[interval]['div_bullish_confirmed'])
                bull_d = len(processed_data[interval]['div_bullish_developing'])
                bull_f = len(processed_data[interval]['div_bullish_forming'])
                bear_c = len(processed_data[interval]['div_bearish_confirmed'])
                bear_d = len(processed_data[interval]['div_bearish_developing'])
                bear_f = len(processed_data[interval]['div_bearish_forming'])
                print(f"   🟢 Bullish Div: C={bull_c} D={bull_d} F={bull_f}")
                print(f"   🔴 Bearish Div: C={bear_c} D={bear_d} F={bear_f}")
        print(f"{'='*70}\n")
        
        self._save_to_excel(processed_data)
        self._upload_to_google_sheet(processed_data)
        self._send_telegram_message(processed_data)
        
        print(f'\n🔥 Complete!')


if __name__ == "__main__":
    analyzer = BinanceRSIAnalyzer()
    analyzer.analyze()
