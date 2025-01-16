import pandas as pd
import re
from openpyxl import load_workbook
from datetime import datetime
from bs4 import BeautifulSoup

# File paths
excel_file = './data/home-prices-change.xlsx'
html_template = './home-prices-change/index.html'
output_html = './home-prices-change/index.html'

# Read Excel data
data = pd.read_excel(excel_file, sheet_name="Data", engine="openpyxl")
data["Date"] = pd.to_datetime(data["Date"])
data.sort_values(by="Date", inplace=True)

# Calculate the latest values
latest_date = data.iloc[-1]["Date"]
latest_volume = data.iloc[-1]["Value"]

# Prepare "let pi" data
dates_since_reference = (data["Date"] - datetime(1969, 12, 20)).dt.days.tolist()
monthly_totals = data["Value"].tolist()
pi_data = f"let pi = [{dates_since_reference}, {monthly_totals}, null, null, '%', 0, []];"

# Read the HTML template
with open(html_template, "r", encoding="utf-8") as file:
    html_content = file.read()

# Parse the HTML using BeautifulSoup
soup = BeautifulSoup(html_content, "html.parser")

# Update the "let pi" data
script_tags = soup.find_all("script")
for script in script_tags:
    if "let pi = [" in script.text:
        script.string = pi_data
        break

# Update the current values
html_content = re.sub(
    r"<b>Current <span class=\"currentTitle\">.*?</span>:</b>.*?\(.*?\)",
    f"<b>Current <span class=\"currentTitle\">Canada Home Prices Change</span>:</b> {latest_volume:,}%",
    html_content,
    flags=re.DOTALL
)

# Update the timestamp
timestamp_div = soup.find("div", {"id": "timestamp"})
if timestamp_div:
    timestamp_div.string = latest_date.strftime("%b %Y")

# Update the historical average in the stats table
historical_avg_td = soup.select_one("table#stats td.left + td")
if historical_avg_td:
    historical_avg_td.string = "3%"

# Save the updated HTML
with open(output_html, "w", encoding="utf-8") as file:
    file.write(str(soup))

print("HTML updated successfully!")
