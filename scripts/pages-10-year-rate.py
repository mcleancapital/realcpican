import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import re

# Paths for the files
excel_file = './data/10-year-treasury-rate.xlsx'
html_template = './10-year-treasury-rate/index.html'
output_html = './10-year-treasury-rate/index.html'

# Read Excel data
data = pd.read_excel(excel_file, sheet_name="Data", engine="openpyxl")
data["Date"] = pd.to_datetime(data["Date"])
data.sort_values(by="Date", inplace=True)

# Calculate the latest values
latest_date = data.iloc[-1]["Date"]
latest_volume = data.iloc[-1]["Value"]

# Read cell C2 for '% Change vs Last Year' using openpyxl
try:
    wb = load_workbook(excel_file, data_only=True)  # data_only=True retrieves evaluated formula values
    sheet = wb["Data"]
    latest_percentage_change = sheet["C2"].value  # Retrieve the value of cell C2
    print(f"Value of C2 (openpyxl): {latest_percentage_change}")
except Exception as e:
    print(f"Error reading C2 with openpyxl: {e}")
    latest_percentage_change = None

# Handle cases where C2 is None or invalid
formatted_percentage_change = f"{latest_percentage_change:+.1f}%" if latest_percentage_change else "N/A"

# Format the "let pi" data
dates_since_reference = (data["Date"] - datetime(1970, 1, 1)).dt.days.tolist()
monthly_totals = data["Value"].tolist()
pi_data = f"let pi = [{dates_since_reference}, {monthly_totals}, null, null, '%', 0, []];"

# Read the HTML template
with open(html_template, "r", encoding="utf-8") as file:
    html_content = file.read()

# Replace "let pi" data
html_content = re.sub(r"let pi = \[.*?\];", pi_data, html_content, flags=re.DOTALL)

# Replace the values and percentage change
html_content = re.sub(
    r"<b>Current <span class=\"currentTitle\">.*?</span>:</b>.*?\(.*?\)",
    f"<b>Current <span class=\"currentTitle\">10 Year Treasury Rate</span>:</b> {latest_volume:.2f}% ({formatted_percentage_change} vs last year)",
    html_content,
    flags=re.DOTALL
)

# Update the timestamp
html_content = re.sub(
    r"<div id=\"timestamp\">.*?</div>",
    f"<div id=\"timestamp\">{latest_date.strftime('%b %d, %Y')}</div>",
    html_content,
    flags=re.DOTALL
)

# Save the updated HTML locally
with open(output_html, "w", encoding="utf-8") as file:
    file.write(html_content)
