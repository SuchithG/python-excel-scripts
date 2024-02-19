import os
import pandas as pd
import time
from datetime import datetime

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

    # Replace NaN values with 0 and remove columns starting with 'Unnamed:'
    result_df = result_df.loc[:, ~result_df.columns.str.contains('^Unnamed')]

    # Transpose specified columns
    result_df = result_df.melt(id_vars=["Formula", "Resource name", "Date", "Month"],
                               value_vars=["Count", "Setup", "Amend", "Review", "Closure", "4 eye Count"],
                               var_name='Activity Type',
                               value_name='Value')

    # Save the concatenated data frame to a new Excel file
    output_file_path = os.path.join(output_folder_path, "Resources_Daily_Volumes_Data.xlsx")
    result_df.to_excel(output_file_path, index=False)
    print(f"Data concatenated successfully. Output file saved at: {output_folder_path}")
else:
    print("No valid data found. Concatenation skipped.")

end_time = time.time()
execution_time = (end_time - start_time) / 60
print(f"Script execution completed in {execution_time:.2f} minutes")
