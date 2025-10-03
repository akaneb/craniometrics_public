import pandas as pd
import os
import shutil

# Load the Excel file - before running, update this file path to use the correct source template
file_path = r'C:\Users\akane\Desktop\cm_data\LD_laterality.xlsx'
data = pd.read_excel(file_path)

# Define source and destination directories - update as needed
source_dir = r"C:\Users\akane\Desktop\cm_data\LD_points_parsed"
destination_dir = r"C:\Users\akane\Desktop\cm_data\LD_right"

# Create destination directory if it doesn't exist - likely unnecessary 
os.makedirs(destination_dir, exist_ok=True)

# Function to copy files based on Excel data
def copy_files_based_on_excel(data, source_dir, destination_dir):
    for index, row in data.iterrows():
        file_path = row.iloc[0].strip()  # Trim any extra spaces
        sign = row.iloc[1].strip()  # Trim any extra spaces
        
        # Extract the file name from the file path (handling backslashes)
        file_name = os.path.basename(file_path.replace('\\', '/')).strip()
        
        # Example: Extracted file_name would be '4478222_Khan_Saim_012919LM'
        # Append the suffix and the .xlsm extension to the file name
        source_file_name = file_name + '_GridAnalysisPoints.xlsm'
        source_file_path = os.path.join(source_dir, source_file_name).strip()
        
        # Debugging statements
        print(f"Processing file: {file_name} with sign: {sign}")
        print(f"Looking for file: {source_file_path}")

        # Check if the sign is '-' and if the file exists in the source directory
        if sign == '+':
            if os.path.exists(source_file_path):
                # Copy the file to the destination directory
                shutil.copy(source_file_path, destination_dir)
                print(f"Copied {source_file_name} to {destination_dir}")
            else:
                print(f"File not found: {source_file_path}")
        else:
            print(f"Skipping file: {file_name} with sign: {sign}")

# Execute the function
copy_files_based_on_excel(data, source_dir, destination_dir)
