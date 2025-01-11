import openpyxl
from datetime import datetime

# Paths for the files
excel_file = './data/sectors.xlsx'
html_template = './sectors_local.html'
output_html = output_html = './sectors.html'

# Open the Excel file
wb = openpyxl.load_workbook(excel_file)
sheet = wb.active

# Fetch returns and the last update date
returns = [sheet[f"D{row}"].value for row in range(2, 14)]
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
html_content = html_content.replace("Dec 31, 2024", last_update_date)

# Replace the returns (with corrected order for Amazon and Meta)
html_content = html_content.replace("2.86%", returns[0])  # Energy
html_content = html_content.replace("-0.97%", returns[1])   # Materials
html_content = html_content.replace("-0.28%", returns[2])  # Industrials
html_content = html_content.replace("-1.23%", returns[3])   # Cons. Discr.
html_content = html_content.replace("1.52%", returns[5])   # Cons. Staples
html_content = html_content.replace("-2.24%", returns[4])   # Health Care
html_content = html_content.replace("-2.15%", returns[6])  # Financials
html_content = html_content.replace("-1.73%", returns[7])   # Technology
html_content = html_content.replace("0.73%", returns[8])   # Communications
html_content = html_content.replace("-0.16%", returns[9])   # Utilities
html_content = html_content.replace("-4.36%", returns[10])  # Real Estate
html_content = html_content.replace("-0.93%", returns[11])  # Total S&P/TSX

# Save the updated HTML locally
with open(output_html, "w", encoding="utf-8") as file:
    file.write(html_content)

print("HTML file updated successfully.")
