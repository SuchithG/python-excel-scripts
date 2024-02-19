from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
import os
import pandas as pd
import time
from datetime import datetime
import openpyxl

start_time = time.time()

# Specify the paths to the folders containing the excel files
folder_paths = [
    r'path1',
    r'path2',
    r'path3',
]

# Define the output folder path
output_folder_path = r"output_path"

# Create output folder if not exists
if not os.path.exists(output_folder_path):
    os.makedirs(output_folder_path)
    print(f"Created output folder at: {output_folder_path}")

# List to store data frames
dfs = []

# Loop through each folder path
for folder_path in folder_paths:

    # Check if folder path exists
    if not os.path.exists(folder_path):
        print(f"Folder path '{folder_path}' does not exist. Skipping...")
        continue

    print(f"Processing files in folder: {folder_path}")

    # Loop through each file in the folder
    for filename in os.listdir(folder_path):
        
        # Check if the file ends with .xlsx and is not a temporary file
        if filename.endswith(".xlsx") and not filename.startswith("~$"):
            
            # Define the full path to the file
            file_path = os.path.join(folder_path, filename)
            
            # Try reading the file
            try:
                df = pd.read_excel(file_path)
                
                # Check if necessary columns are in the file
                if set(["Formula", "Resource name", "Date", "Month"]).issubset(df.columns):
                    
                    # Convert 'Month' column to datetime for comparison, assuming 'Month' as 'MMM-YY' format
                    df['Month'] = pd.to_datetime(df['Month'], format='%b-%y')
                    
                    # Keep only rows where 'Month' is on or after January 2023 and necessary columns are not null
                    comparison_date = datetime(2023, 1, 1)  # January 2023
                    df = df[df['Month'] >= comparison_date]
                    df = df[df[["Formula", "Resource name", "Date", "Month"]].notnull().all(axis=1)]
                    
                    if not df.empty:
                        # Dynamically adjust 'Month' column to reflect the month and year correctly
                        # Format 'Month' column to display as 'MMM-YY'
                        df['Month'] = df['Month'].dt.strftime('%b-%y')
                        
                        dfs.append(df)
                        print(f"Added data from file: {filename}")
                    else:
                        print(f"File {filename} does not contain valid data. Skipping...")
            
            except Exception as e:
                print(f"Could not read file {filename}. Error: {e}")

# After concatenating all data frames
if dfs:
    result_df = pd.concat(dfs, ignore_index=True)

    # Correctly handle 'Date' and 'Actual Date' as dates formatted as mm/dd/yyyy
    result_df['Date'] = pd.to_datetime(result_df['Date'], errors='coerce').dt.strftime('%m/%d/%Y')
    result_df['Actual Date'] = pd.to_datetime(result_df['Actual Date'], errors='coerce').dt.strftime('%m/%d/%Y')

    # Prepare 'Month' column as datetime for formatting in Excel (keep as datetime object for now)
    result_df['Month'] = pd.to_datetime(result_df['Month'], format='%b-%y', errors='coerce')
    result_df['Month'] = result_df['Month'].apply(lambda x: x.replace(day=24))

    # Convert 'Resource Name' to uppercase
    result_df['Resource Name'] = result_df['Resource Name'].str.upper()

    # Melt operation
    result_df = result_df.melt(id_vars=["Formula", "Resource Name", "Date", "Month", "Category", "Work Drivers", "Activity", "Asset Class", "Case #", "Error Count", "Actual Date", "ID number"],
                               value_vars=["Count", "Setup", "Amend", "Review", "Closure", "4 eye Count"],
                               var_name='Name',
                               value_name='Value')

    # Add 'Source' and 'InAccuracy' columns with static values
    result_df['Source'] = "Orchestra"
    result_df['InAccuracy'] = "Accurate"

    # Ensure the columns are in the specified order
    result_df = result_df[["Formula", "Resource Name", "Date", "Month", "Category", "Work Drivers", "Activity", "Asset Class", "Case #", "Error Count", "Actual Date", "ID number", "Source", "Name", "Value", "InAccuracy"]]

    # Save the DataFrame to an Excel file
    output_file_path = os.path.join(output_folder_path, "Resources_Daily_Volumes_Data.xlsx")
    result_df.to_excel(output_file_path, index=False)

    # Now, use openpyxl to apply custom formatting for the 'Month' column
    wb = load_workbook(output_file_path)
    ws = wb.active

    # Define a custom date style for 'Month'
    date_style = NamedStyle(name='date_style', number_format='dd-mmm')
    wb.add_named_style(date_style)

    for cell in ws['D'][1:]:  # Assuming 'Month' is column D; adjust if your DataFrame structure is different
        cell.style = date_style

    wb.save(output_file_path)

    print(f"Data concatenated successfully. Output file saved at: {output_folder_path}")
else:
    print("No valid data found. Concatenation skipped.")

end_time = time.time()
execution_time = (end_time - start_time) / 60
print(f"Script execution completed in {execution_time:.2f} minutes")