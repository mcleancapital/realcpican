import openpyxl
from datetime import datetime
from yfinance import Ticker

# Define the stock symbols and the Excel file path
stock_symbols = ["RY.to", "SHOP.to", "TD.to", "ENB.to", "BNS.to", "BMO.to", "CP.to", "CNQ.to", "TRI.to", "BN.to"]
excel_file_path = './data/top10.xlsx'

# Function to fetch the most recent data (date and price) for a stock
def fetch_recent_stock_data(symbol):
    try:
        ticker = Ticker(symbol)
        hist = ticker.history(period="1mo")  # Fetch the last month's data
        recent_date = hist.index[-1].date()  # Most recent date
        recent_price = hist['Close'][-1]    # Most recent closing price
        return recent_date, recent_price
    except Exception as e:
        print(f"Error fetching data for {symbol}: {e}")
        return None, None

# Open the Excel file
wb = openpyxl.load_workbook(excel_file_path)
sheet = wb.active

# Fetch and write the last closing prices and dates
for idx, symbol in enumerate(stock_symbols, start=1):
    try:
        date, price = fetch_recent_stock_data(symbol)
        if date and price:
            sheet[f"C{idx + 1}"] = price  # Write price in column C
            sheet[f"E{idx + 1}"] = date.strftime('%Y-%m-%d')  # Write date in column E
            print(f"Fetched data for {symbol}: Price = {price}, Date = {date}")
    except Exception as e:
        print(f"Error processing data for {symbol}: {e}")

# Calculate and write the percentage change in column D
for row in range(2, len(stock_symbols) + 2):
    try:
        previous_price = sheet[f"B{row}"].value
        current_price = sheet[f"C{row}"].value

        if previous_price is not None and current_price is not None:
            change = (current_price / previous_price - 1) * 100
            sheet[f"D{row}"] = round(change, 2)  # Write change in column D
    except Exception as e:
        print(f"Error calculating change for row {row}: {e}")

# Save the updated Excel file
wb.save(excel_file_path)
print(f"Updated Excel file saved at {excel_file_path}")
