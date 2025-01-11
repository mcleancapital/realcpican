import openpyxl
from datetime import datetime

def update_sp500_html(html_file, excel_file, output_file):
    try:
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

        print("Step 2: Reading HTML file...")
        # Read the HTML content
        with open(html_file, "r", encoding="utf-8") as file:
            html_content = file.read()

        # Step 3: Locate the specific section for TOP 10
        print("Step 4: Updating the specific section for TOP 10...")
        sp500_marker = '<a class=box href="/top10">'
        marker_start = html_content.find(sp500_marker)
        if marker_start == -1:
            print("Marker for MAG7 not found in the HTML.")
            return

        # Locate the end of this section
        section_end = html_content.find("</a>", marker_start) + 4
        section_content = html_content[marker_start:section_end]

        # Update the value <div> within this section
        value_start = section_content.find("<div>", section_content.find("<h3>")) + 5
        value_end = section_content.find("</div>", value_start)

        # Combine the updated value and change
        updated_value = f"{total_average_str}"

        # Update the section with the new value
        updated_section = (
            section_content[:value_start] +
            updated_value +
            section_content[value_end:]
        )

        # Update the date <div> within this section
        date_marker = '<div class="date">'
        date_start = updated_section.find(date_marker) + len(date_marker)
        date_end = updated_section.find("</div>", date_start)
        updated_section = (
            updated_section[:date_start] +
            last_update_date +
            updated_section[date_end:]
        )
        
        # Replace the original section in the HTML
        html_content = (
            html_content[:marker_start] +
            updated_section +
            html_content[section_end:]
        )

        # Step 4: Save the updated HTML
        print("Step 5: Writing updated HTML to output file...")
        with open(output_file, "w", encoding="utf-8") as file:
            file.write(html_content)

        print(f"HTML file '{output_file}' updated successfully!")

    except Exception as e:
        print(f"An error occurred: {e}")

# File paths
html_file = './index.html'
excel_file = './data/top10.xlsx'
output_file = './index.html'

# Run the update function
update_sp500_html(html_file, excel_file, output_file)
