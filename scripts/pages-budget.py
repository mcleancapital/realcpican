import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import re

# Paths for the files
excel_file = './data/budget.xlsx'
html_template = './budget/index.html'
output_html = './budget/index.html'

# Read Excel data
data = pd.read_excel(excel_file, sheet_name="Data", engine="openpyxl")
data["Date"] = pd.to_datetime(data["Date"])
data.sort_values(by="Date", inplace=True)

# Calculate the latest values
latest_date = data.iloc[-1]["Date"]
latest_volume = data.iloc[-1]["Value"]

# Format the "let pi" data
dates_since_reference = (data["Date"] - datetime(1970, 1, 1)).dt.days.tolist()
monthly_totals = data["Value"].tolist()
pi_data = f"let pi = [{dates_since_reference}, {monthly_totals}, null, null, '%', 0, []];"

# Read the HTML template
with open(html_template, "r", encoding="utf-8") as file:
    html_content = file.read()

# Replace "let pi" data
html_content = re.sub(
    r"let pi = \[.*?\];", 
    pi_data, 
    html_content, 
    flags=re.DOTALL
)

# Replace the current value and timestamp
html_content = re.sub(
    r"<b>Current <span class=\"currentTitle\">.*?</span>:</b>.*?<div id=\"timestamp\">.*?</div>",
    f"<b>Current <span class=\"currentTitle\">Canada Government Budget (surplus or deficit as a % of GDP)</span>:</b> {latest_volume:.2f}%"
    f"<div id=\"timestamp\">{latest_date.strftime('%b %Y')}</div>",
    html_content,
    flags=re.DOTALL
)

# Save the updated HTML locally
with open(output_html, "w", encoding="utf-8") as file:
    file.write(html_content)
