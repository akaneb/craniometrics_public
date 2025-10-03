## This is the FIRST STEP in the craniometrics database workflow when you have new data FROM 3D MD
## Follow the green comments for areas to edit before running
## To run, open PowerShell from your computer, run the command 'cd "insert\working\directory\here"
## Then, run 'python data_run.py'

import pandas as pd
import openpyxl
import os
import csv
import win32com.client as win32

# Path to the Excel template - update depending on the population
template_path = r'C:\Users\akane\Desktop\cm_data\LDI_template.xlsm'

# Directory containing the CSV files - update depending on population
csv_directory = r'C:\Users\akane\Desktop\cm_data\sample_directory'

# Directory to save processed Excel files - update depending on populations
output_directory = r'C:\Users\akane\Desktop\cm_data\output_sample'
os.makedirs(output_directory, exist_ok=True)

# Function to read CSV file using the csv module - do not update
def read_csv_file(file_path):
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        data = list(reader)
    return data

# Function to process each CSV file - do not update
def process_csv(csv_file):
    try:
        # Read the CSV file - do not update
        data = read_csv_file(csv_file)
    except Exception as e:
        print(f"Error reading {csv_file}: {e}")
        return

    # Open Excel with COM object - do not update
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    workbook = excel.Workbooks.Open(template_path)
    sheet = workbook.Sheets('unproc_data')

    try:
        # Clear existing data in the sheet - do not update
        used_range = sheet.UsedRange
        if used_range:
            sheet.Range("A2:Z" + str(used_range.Rows.Count)).ClearContents()
        else:
            print(f"No used range found in the sheet for file: {csv_file}")
    except Exception as e:
        print(f"Error clearing contents for file {csv_file}: {e}")
        workbook.Close(SaveChanges=False)
        excel.Application.Quit()
        return

    # Write the CSV data to the sheet - do not update
    for row_index, row in enumerate(data, start=1):
        for col_index, value in enumerate(row, start=1):
            # Strip leading/trailing whitespace from each value - do not update
            sheet.Cells(row_index, col_index).Value = value.strip() if isinstance(value, str) else value

    # Save the processed Excel file - do not update
    output_path = os.path.join(output_directory, os.path.basename(csv_file).replace('.csv', '.xlsm'))
    workbook.SaveAs(output_path, FileFormat=52)  # 52 is the file format code for .xlsm
    workbook.Close(SaveChanges=True)
    excel.Application.Quit()
    print(f'Processed and saved: {output_path}')

# Loop through all CSV files in the directory - do not update
for csv_file in os.listdir(csv_directory):
    if csv_file.endswith('.csv'):
        process_csv(os.path.join(csv_directory, csv_file))

print('All files processed.')
