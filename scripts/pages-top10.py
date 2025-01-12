import openpyxl
from datetime import datetime

# Paths for the files
excel_file = './data/top10.xlsx'
html_template = './top10/index_local.html'
output_html = output_html = './top10/index.html'

# Open the Excel file
wb = openpyxl.load_workbook(excel_file)
sheet = wb.active

# Fetch returns and the last update date
returns = [sheet[f"D{row}"].value for row in range(2, 12)]
raw_date = sheet["E2"].value

# Parse the date properly
if isinstance(raw_date, datetime):
    last_update_date = raw_date.strftime("%b %d, %Y")
elif isinstance(raw_date, str):
    try:
        parsed_date = datetime.strptime(raw_date, "%Y-%m-%d")
        last_update_date = parsed_date.strftime("%b %d, %Y")
    except ValueError:
        last_update_date = "N/A"
else:
    last_update_date = "N/A"

# Calculate the average return for "Total"
valid_returns = [ret for ret in returns if isinstance(ret, (int, float))]

if valid_returns:
    total_average = sum(valid_returns) / len(valid_returns)
else:
    total_average = 0

# Ensure all data is valid
returns = [f"{ret:.2f}%" if isinstance(ret, (int, float)) else "N/A" for ret in returns]
total_average_str = f"{total_average:.2f}%"

# Read the HTML template
with open(html_template, "r", encoding="utf-8") as file:
    html_content = file.read()

# Replace the last update date
html_content = html_content.replace("Jan 7, 2025", last_update_date)

# Replace the returns (with corrected order for Amazon and Meta)
html_content = html_content.replace("-4.3%", returns[0])  # Royal Bank
html_content = html_content.replace("1.9%", returns[1])   # Shopify
html_content = html_content.replace("-0.1%", returns[2])  # TD
html_content = html_content.replace("1.4%", returns[3])   # Enbridge
html_content = html_content.replace("0.4%", returns[5])   # BMO
html_content = html_content.replace("3.3%", returns[4])   # Bank of Nova Scotia
html_content = html_content.replace("-5.4%", returns[6])  # Canadian Pacific Railway
html_content = html_content.replace("0.2%", returns[7])   # Canadian Natural Resources
html_content = html_content.replace("0.3%", returns[8])   # Thomson Reuters
html_content = html_content.replace("0.5%", returns[9])  # Brookfield

# Replace the total return
html_content = html_content.replace("10.0%", total_average_str)

# Save the updated HTML locally
with open(output_html, "w", encoding="utf-8") as file:
    file.write(html_content)

print("HTML file updated successfully.")
