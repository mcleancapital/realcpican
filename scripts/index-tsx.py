import pandas as pd
from datetime import datetime

def update_sp500_html(html_file, excel_file, output_file):
    try:
        print("Step 1: Reading Excel file...")
        # Read the Excel file
        df = pd.read_excel(excel_file, sheet_name="Data", usecols=["Date", "Value"], header=0)

        # Drop rows where 'Date' or 'Value' is missing
        df = df.dropna(subset=["Date", "Value"])

        # Convert 'Date' to datetime
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date"])  # Remove rows with invalid dates

        # Calculate numeric representation of dates relative to 1970-01-01
        epoch = datetime(1970, 1, 1)
        df["Date_Numeric"] = (df["Date"] - epoch).dt.days

        # Sort data in ascending order of dates
        df = df.sort_values(by="Date_Numeric", ascending=True).reset_index(drop=True)

        # Extract arrays for the output
        date_array = df["Date_Numeric"].tolist()
        value_array = df["Value"].tolist()

        # Format arrays as strings
        formatted_dates = ", ".join(map(str, date_array))
        formatted_values = ", ".join(map(str, value_array))

        # Combine the formatted array with the required suffix
        formatted_data = f"[[{formatted_dates}], [{formatted_values}], null, null, '', 1, []]"

        # Get the most recent date and value
        most_recent_date = df.iloc[-1]["Date"]
        most_recent_value = df.iloc[-1]["Value"]

        # Format the date into "4:00 PM EST, Fri Dec 13" format
        formatted_date = most_recent_date.strftime("4:00 PM EST, %a %b %d")
        formatted_value = f"{most_recent_value:,.2f}"

        print("Step 2: Reading HTML file...")
        # Read the HTML content
        with open(html_file, "r", encoding="utf-8") as file:
            html_content = file.read()

        # Step 3: Update the data section in HTML
        print("Step 3: Updating the data section in HTML...")
        data_marker = "<!-- TSX historical prices -->"
        if data_marker in html_content:
            data_start = html_content.find(data_marker) + len(data_marker)
            data_end = html_content.find("]]", data_start) + 2  # Locate the end of the array
            html_content = (
                html_content[:data_start] +
                "\n" +
                formatted_data +
                "\n" +
                html_content[data_end:]
            )
        else:
            print(f"Data section marker '{data_marker}' not found in HTML.")
            return

        # Step 4: Locate the specific section for S&P 500 Historical Prices
        print("Step 4: Updating the specific section for S&P 500 Historical Prices...")
        sp500_marker = '<a class=box href="/tsx-historical-prices">'
        marker_start = html_content.find(sp500_marker)
        if marker_start == -1:
            print("Marker for S&P 500 Historical Prices not found in the HTML.")
            return

        # Locate the end of this section
        section_end = html_content.find("</a>", marker_start) + 4
        section_content = html_content[marker_start:section_end]

        # Update the value <div> within this section
        value_start = section_content.find("<div>", section_content.find("<h3>")) + 5
        value_end = section_content.find("</div>", value_start)
        updated_section = (
            section_content[:value_start] +
            formatted_value +
            section_content[value_end:]
        )

        # Update the date <div> within this section
        date_marker = '<div class="date">'
        date_start = updated_section.find(date_marker) + len(date_marker)
        date_end = updated_section.find("</div>", date_start)
        updated_section = (
            updated_section[:date_start] +
            formatted_date +
            updated_section[date_end:]
        )

        # Replace the original section in the HTML
        html_content = (
            html_content[:marker_start] +
            updated_section +
            html_content[section_end:]
        )

        # Step 5: Save the updated HTML
        print("Step 5: Writing updated HTML to output file...")
        with open(output_file, "w", encoding="utf-8") as file:
            file.write(html_content)

        print(f"HTML file '{output_file}' updated successfully!")

    except Exception as e:
        print(f"An error occurred: {e}")

# File paths
html_file = './index.html'
excel_file = './data/tsx.xlsx'
output_file = './index.html'

# Run the update function
update_sp500_html(html_file, excel_file, output_file)
