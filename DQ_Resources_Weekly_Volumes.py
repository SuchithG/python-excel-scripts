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

# Function to check if the month is equal to or greater than Jan-23
def month_ge_jan_23(month):
    comparison_date = datetime.strptime("Jan-23", "%b-%y")
    month_date = datetime.strptime(month, "%b-%y")
    return month_date >= comparison_date

# Create output folder if not exists
if not os.path.exists(output_folder_path):
    os.makedirs(output_folder_path)
    print(f"Created output folder at: {output_folder_path}")

# List to store data frames
dfs = []

# Loop through each folder path
for folder_path in folder_paths:
    if not os.path.exists(folder_path):
        print(f"Folder path '{folder_path}' does not exist. Skipping...")
        continue

    print(f"Processing files in folder: {folder_path}")
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx") and not filename.startswith("~$"):
            file_path = os.path.join(folder_path, filename)
            try:
                df = pd.read_excel(file_path)
                if set(["Formula", "Resource name", "Date", "Month"]).issubset(df.columns):
                    df['Date'] = pd.to_datetime(df['Date'])
                    df['Month'] = pd.to_datetime(df['Month'], format='%b-%y')
                    
                    # Filter based on the 'Month' condition
                    df = df[df['Month'].apply(lambda x: month_ge_jan_23(x.strftime('%b-%y')))]
                    df['Date'] = df['Date'].dt.strftime('%d-%b')  # Convert date format
                    
                    if not df.empty:
                        dfs.append(df)
                        print(f"Added data from file: {filename}")
                    else:
                        print(f"File {filename} does not contain data for Jan-23 or later. Skipping...")
            except Exception as e:
                print(f"Could not process file {filename}. Error: {e}")

if dfs:
    result_df = pd.concat(dfs, ignore_index=True)
    result_df = result_df.loc[:, ~result_df.columns.str.contains('Unnamed:')]
    result_df = result_df.melt(id_vars=["Formula", "Resource name", "Date", "Month", "Category", "Work Drivers", "Activity", "Asset Class", "Case #", "Error Count", "Actual Date", "ID number"],
                               value_vars=["Count", "Setup", "Amend", "Review", "Closure", "4 eye Count"],
                               var_name='Activity Type',
                               value_name='Value')

    output_file_path = os.path.join(output_folder_path, "Resources_Daily_Volumes_Data.xlsx")
    result_df.to_excel(output_file_path, index=False)
    print(f"Data concatenated successfully. Output file saved at: {output_file_path}")
else:
    print("No valid data found. Concatenation skipped.")

end_time = time.time()
execution_time_in_minutes = (end_time - start_time) / 60
print(f"Script execution completed in {execution_time_in_minutes:.2f} minutes")
