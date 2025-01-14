# PDF Invoice Data Extraction Tool

This tool is designed to efficiently extract key data from invoice PDFs, such as dates and labeled values, and save the results into structured Excel and CSV files. It simplifies the process of extracting invoice details, making it useful for businesses looking to automate and streamline invoice processing.

## Features

- **Extracts Key Data**: Extracts information like invoice dates (e.g., "Invoice Date") and associated values (e.g., "Gross Amount incl. VAT", "Total").
- **Supports Multiple Formats**: Handles both table-based and text-based invoice formats. It can process dates in German and English formats.
- **Data Output**: Saves extracted data in two formats for easy analysis:
  - **Excel File** (`Invoice_Data.xlsx`):
    - **Sheet 1**: Contains three columns: "File Name", "Date", and "Value".
    - **Sheet 2**: Includes a pivot table that summarizes data by "Date", "Value", and "File Name".
  - **CSV File** (`Invoice_Data.csv`): A simple and structured CSV file containing all the extracted data.

## How to Use

1. Download the script.zip file containing the executable and necessary PDFs from the [First Release](https://github.com/SripalUdyavara/Udyavara_Sripal_WerkStudent_Python/releases/tag/v1.0.0) section of the GitHub repository.
2. Extract the contents of the zip file to a folder on your computer.
3. Run the executable file:
   - On **Windows**: Double-click `script.exe`.
   - On **Mac/Linux**: Open the terminal, navigate to the folder containing the tool, and run:
     ```bash
     ./script
     ```
4. After running the tool:
   - Open `Invoice_Data.xlsx` or `Invoice_Data.csv` in the same folder to view the extracted results.

## Technical Overview

1. **PDF Processing**: The tool reads PDF invoices, extracting data associated with labels like "Invoice Date" and values such as "Gross Amount incl. VAT".
2. **Flexible Date Handling**: It recognizes both German (e.g., "1. März 2024") and English (e.g., "Nov 26, 2016") date formats.
3. **Easy-to-Use Output**:
   - **Excel File**: Provides two sheets—one with raw data and the other with a pivot table summary for easy analysis.
   - **CSV File**: Generates a straightforward CSV file for further data manipulation or export.

## System Requirements

- **Python 3.6+** (optional if running the Python script).
- The following Python libraries must be installed if running the Python script:
  - `pdfplumber`
  - `pandas`
  - `openpyxl`

  You can install the necessary dependencies using:
  ```bash
  pip install -r requirements.txt
  ```
