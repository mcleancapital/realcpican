import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import re

# Paths for the files
excel_file = './data/wages.xlsx'
html_template = './wages/index.html'
output_html = './wages/index.html'

# Read Excel data
data = pd.read_excel(excel_file, sheet_name="Data", engine="openpyxl")
data["Date"] = pd.to_datetime(data["Date"])
data.sort_values(by="Date", inplace=True)

# Calculate the latest values
latest_date = data.iloc[-1]["Date"]
latest_volume = data.iloc[-1]["Value"]

# Format the "let pi" data
dates_since_reference = (data["Date"] - datetime(1969, 12, 20)).dt.days.tolist()
monthly_totals = data["Value"].tolist()
pi_data = f"let pi = [{dates_since_reference}, {monthly_totals}, null, null, '%', 0, []];"

# Read the HTML template
with open(html_template, "r", encoding="utf-8") as file:
    html_content = file.read()

# Replace "let pi" data
html_content = re.sub(r"let pi = \[.*?\];", pi_data, html_content, flags=re.DOTALL)

# Replace the values
html_content = re.sub(
    r"<b>Current <span class=\"currentTitle\">.*?</span>:</b>.*?\(.*?\)",
    f"<b>Current <span class=\"currentTitle\">Canada Wages Growth</span>:</b> {latest_volume:,}%",
    html_content,
    flags=re.DOTALL
)

# Update the timestamp
html_content = re.sub(
    r"<div id=\"timestamp\">.*?</div>",
    f"<div id=\"timestamp\">{latest_date.strftime('%b %Y')}</div>",
    html_content,
    flags=re.DOTALL
)

# Save the updated HTML locally
with open(output_html, "w", encoding="utf-8") as file:
    file.write(html_content)
