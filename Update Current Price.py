import openpyxl
import os
from datetime import datetime
import yfinance as yf

excel_file_path = "Portfolio Tracker\\User.xlsx"
backup_directory = "Portfolio Tracker\\Records"
sheet_name = 'Sheet1'

def get_stock_price(symbol):
    try:
        ticker = yf.Ticker(symbol)
        historical_data = ticker.history(period="1d")
        price = historical_data['Close'].iloc[-1]
        return float(price)
    except Exception as e:
        print(f"Error fetching price for {symbol}: {e}")
        return None

def write_to_excel(file_path, cell, data):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        sheet[cell] = data
        workbook.save(file_path)
        return data
    except Exception as e:
        print(f"Error writing to Excel file: {e}")
        return None
        
def backup_excel_file(file_path, backup_directory):
    try:
        timestamp = datetime.now().strftime("%d_%m_%Y")
        backup_file_name = f"Record_Portfolio_{timestamp}.xlsx"
        backup_file_path = os.path.join(backup_directory, backup_file_name)
        os.system(f'copy "{file_path}" "{backup_file_path}"')
        print(f"Backup created: {backup_file_path}")
    except Exception as e:
        print(f"Error creating backup: {e}")
        
def read_excel_cell(file_path, sheet_name, cell):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        cell_value = sheet[cell].value
        return cell_value
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None
    
quantity_col = "B"
buying_price_col = "C"
stock_name_col = "A"
current_price_col = "E"

for i in range(2, 9):
    try:
        # Read stock symbol and name
        stock_name_cell = stock_name_col + str(i)
        stock_name = read_excel_cell(excel_file_path, sheet_name, stock_name_cell)
        
        stock_symbol_cell = "K" + str(i)
        stock_symbol = read_excel_cell(excel_file_path, sheet_name, stock_symbol_cell)
        
        # Get current stock price
        current_price = get_stock_price(stock_symbol)
        if current_price is None:
            print(f"Skipping {stock_symbol} due to price fetch error.")
            continue
        
        # Write current price to Excel
        current_price_cell = current_price_col + str(i)
        write_to_excel(excel_file_path, current_price_cell, current_price)
        
        # Read quantity and buying price
        quantity_cell = quantity_col + str(i)
        quantity = int(read_excel_cell(excel_file_path, sheet_name, quantity_cell))
        
        buying_price_cell = buying_price_col + str(i)
        buying_price = float(read_excel_cell(excel_file_path, sheet_name, buying_price_cell))
        
        # Calculate investment and outcome
        invested = quantity * buying_price
        outcome = round(quantity * current_price, 2)
        outcome_diff = round(outcome - invested, 2)
        
        # Print profit or loss statement
        if outcome_diff > 0:
            print(f"PROFIT ~ {stock_name} ~ {quantity} shares ~ ${invested:.2f} ~ ${outcome:.2f} ~ ${outcome_diff:.2f}")
        else:
            print(f"LOSS ~ {stock_name} ~ {quantity} shares ~ ${invested:.2f} ~ ${outcome:.2f} ~ $({-outcome_diff:.2f})")
    
    except Exception as e:
        print(f"Error processing row {i}: {e}")

print()
backup_excel_file(excel_file_path, backup_directory)
