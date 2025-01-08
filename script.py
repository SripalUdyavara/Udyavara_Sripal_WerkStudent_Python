import pdfplumber
import pandas as pd
import re
from datetime import datetime
from openpyxl.utils.dataframe import dataframe_to_rows


# Utility Functions

def extract_date_from_table(pdf_path, date_label):
    """
    Extracts a German-style date from a table where the date is below the label.

    Parameters:
        pdf_path (str): Path to the PDF file.
        date_label (str): The label to search for in the table header (e.g., 'Date').

    Returns:
        str: The extracted date in DD.MM.YYYY format, or 'Date not found' if no date is found.
    """
    # Mapping German month names to their numeric equivalents
    german_months = {
        "Januar": "01", "Februar": "02", "März": "03", "April": "04", "Mai": "05", "Juni": "06",
        "Juli": "07", "August": "08", "September": "09", "Oktober": "10", "November": "11", "Dezember": "12"
    }

    # Regex pattern to match German-style dates
    date_pattern = r"(\d{1,2})\.\s?(\w+)\s(\d{4})"

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables:
                continue

            for table in tables:
                headers = table[0]  # Assuming the first row contains headers
                if date_label in headers:
                    date_col_index = headers.index(date_label)
                    for row in table[1:]:  # Skip the header row
                        cell = row[date_col_index] if len(row) > date_col_index else None
                        if cell and re.search(date_pattern, cell):
                            match = re.search(date_pattern, cell)
                            if match:
                                return format_german_date(match, german_months)

    return "Date not found"


def extract_date_from_text(pdf_path, date_label):
    """
    Extracts a date from plain text where the date label is next to the value.

    Parameters:
        pdf_path (str): Path to the PDF file.
        date_label (str): The label to search for in the text (e.g., 'Invoice date').

    Returns:
        str: The extracted date in DD.MM.YYYY format, or 'Date not found' if no date is found.
    """
    # Regex pattern for English-style dates (e.g., 'Nov 26, 2016')
    date_pattern = r"\b\w{3,9}\s\d{1,2},\s\d{4}"

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            lines = page_text.split('\n')

            for line in lines:
                if date_label in line:  # Search for the label in the line
                    print(f"[DEBUG] Found label '{date_label}' in line: {line}")
                    text_after_label = line.split(date_label, 1)[-1].strip()  # Extract text after the label
                    match = re.search(date_pattern, text_after_label)
                    if match:
                        return format_english_date(match.group())

    return "Date not found"


def extract_value_from_text(pdf_path, keyword):
    """
    Extracts the value associated with a specific keyword from the text.

    Parameters:
        pdf_path (str): Path to the PDF file.
        keyword (str): The keyword to search for in the text (e.g., 'Total').

    Returns:
        str: The extracted value, or 'Value not found' if the keyword is not found.
    """
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            lines = page_text.split('\n')

            for line in lines:
                if keyword in line:  # Search for the keyword in the line
                    print(f"[DEBUG] Found keyword '{keyword}' in line: {line}")
                    value = line.split(keyword, 1)[-1].strip()  # Extract the value after the keyword
                    return value

    return "Value not found"


def format_german_date(match, german_months):
    """
    Converts a German date string into DD.MM.YYYY format.

    Parameters:
        match (re.Match): Regex match object containing the day, month, and year components.
        german_months (dict): Dictionary mapping German month names to numeric values.

    Returns:
        str: The normalized date in DD.MM.YYYY format.
    """
    day, month, year = match.groups()
    if month in german_months:
        return f"{day.zfill(2)}.{german_months[month]}.{year}"
    return "Date not found"


def format_english_date(date_str):
    """
    Converts an English-style date string into DD.MM.YYYY format.

    Parameters:
        date_str (str): The date string to normalize.

    Returns:
        str: The normalized date in DD.MM.YYYY format, or 'Date not found' if normalization fails.
    """
    try:
        # Parse the English-style date (e.g., 'Nov 26, 2016') into a datetime object
        date_obj = datetime.strptime(date_str, "%b %d, %Y")
        return date_obj.strftime("%d.%m.%Y")  # Format the date as DD.MM.YYYY
    except ValueError as e:
        print(f"[DEBUG] Error normalizing date '{date_str}': {e}")

    return "Date not found"


# Main Script

# Define the input PDF files and their configurations
pdf_configurations = [
    {
        "file_name": "sample_invoice_1.pdf",
        "date_label": "Date",
        "is_table": True,
        "additional_labels": ["Gross Amount incl. VAT"],  # Add additional labels as needed
    },
    {
        "file_name": "sample_invoice_2.pdf",
        "date_label": "Invoice date",
        "is_table": False,
        "additional_labels": ["Total"],  # Add additional labels as needed
    },
]

# Extract data from the PDFs and store it in a structured format
extracted_data = []
for pdf_config in pdf_configurations:
    file_name = pdf_config["file_name"]
    date_label = pdf_config["date_label"]
    is_table = pdf_config["is_table"]
    additional_labels = pdf_config["additional_labels"]

    print(f"\n--- Processing File: {file_name} ---\n")
    if is_table:
        extracted_date = extract_date_from_table(file_name, date_label)
    else:
        extracted_date = extract_date_from_text(file_name, date_label)

    # Extract additional values based on keywords
    additional_values = {label: extract_value_from_text(file_name, label) for label in additional_labels}

    # Combine all extracted data into a single row under the 'Value' column
    for label, value in additional_values.items():
        row = {"File Name": file_name, "Extracted Date": extracted_date, "Value": value}
        extracted_data.append(row)

# Convert extracted data into a DataFrame
data_frame = pd.DataFrame(extracted_data)

# Define the output file paths
excel_file_path = "Invoice_Data.xlsx"
csv_file_path = "Invoice_Data.csv"

# Save DataFrame to Excel and CSV
with pd.ExcelWriter(excel_file_path, engine="openpyxl") as writer:
    data_frame.to_excel(writer, sheet_name="Data", index=False)

    # Access the workbook and the data sheet
    workbook = writer.book
    sheet = workbook["Data"]

    # Create a Pivot Table (manually using pandas)
    pivot_data = pd.pivot_table(data_frame, values='Value', index='Extracted Date', columns='File Name', aggfunc='sum')
    pivot_data.reset_index(inplace=True)  # To ensure the pivot data is properly aligned with column headers

    # Add the pivot data to a new sheet
    pivot_sheet = workbook.create_sheet("Pivot")
    for row in dataframe_to_rows(pivot_data, index=False, header=True):
        pivot_sheet.append(row)

# Save the workbook
workbook.save(excel_file_path)

# Save DataFrame to CSV
data_frame.to_csv(csv_file_path, sep=";", index=False)

# Print confirmation messages
print("\nData extraction complete!")
print(f"Excel file created: {excel_file_path}")
print(f"CSV file created: {csv_file_path}")
