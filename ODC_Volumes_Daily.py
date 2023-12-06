import os
import pandas as pd
import time
from datetime import datetime, timedelta

start_time = time.time()

folder_path = "path1"

output_folder_path = "output_path"

# Function to get the previous working day's date in the format 'mm/d/yyyy'
def get_previous_working_day():
    today = datetime.now()
    one_day = timedelta(days=1)
    while True:
        today -= one_day
        if today.weekday() < 5:  # Monday to Friday (0 to 4)
            return today.strftime('%m/%d/%Y')

previous_working_day = get_previous_working_day()
print(f"Filtering data for the previous working day: {previous_working_day}")

if not os.path.exists(output_folder_path):
    os.makedirs(output_folder_path)
    print(f"Created output folder at: {output_folder_path}")

dfs = []

# Check if folder path exists
if not os.path.exists(folder_path):
    print(f"Folder path '{folder_path}' does not exist. Skipping...")
else:
    print(f"Processing files in folder: {folder_path}")

    # Loop through each file in the folder
    for filename in os.listdir(folder_path):
        
        # Check if the file ends with .xlsm.xlsx and is not a temporary file
        if filename.endswith(".xlsm.xlsx") and not filename.startswith("~$"):
            
            # Define the full path to the file
            file_path = os.path.join(folder_path, filename)
            
            # Try reading the file with header
            try:
                df = pd.read_excel(file_path)
                
                # Check if necessary columns are in the file
                if set(["Formula", "Resource name", "Date", "Month"]).issubset(df.columns):

                    # Convert 'Date' and 'Actual Date to upload' columns to datetime
                    df['Date'] = pd.to_datetime(df['Date']).apply(lambda x: f"{x.month}/{x.day}/{x.year}")
                    df['Actual Date of upload'] = pd.to_datetime(df['Actual Date of upload']).dt.strftime("%Y-%m-%d")
                    
                    # Keep only rows where necessary columns are not null
                    df = df[df[["Formula", "Resource name", "Date", "Month"]].notnull().all(axis=1) & (df["Month"] == get_previous_working_day)]
                    
                    if not df.empty:
                        dfs.append(df)
                        print(f"Added data from file: {filename}")
                    else:
                        print(f"File {filename} does not contain valid data. Skipping...")
            
            except Exception as e:
                print(f"Could not read file {filename} with header. Error: {e}")
                try:
                    pd.read_excel(file_path, header=None)
                    print(f"File {filename} read without header. Skipping...")
                except Exception as e:
                    print(f"Could not read file {filename} without header. Error: {e}")
                    pass

# After concatenating all data frames
if dfs:
    result_df = pd.concat(dfs, ignore_index=True)

    # Replace NaN values with 0
    result_df = result_df.loc[:, ~result_df.columns.str.startswith('Unnamed: ')]

    # Save the concatenated data frame to a new Excel file
    output_file_path = os.path.join(output_folder_path, "concatenated_data.xlsx")
    result_df.to_excel(output_file_path, index=False)
    print(f"Data concatenated successfully. Output file saved at: {output_file_path}")
else:
    print("No valid data found. Concatenation skipped.")

end_time = time.time()

execution_time_in_minutes = (end_time - start_time) / 60
print(f"Script execution in {execution_time_in_minutes:.2f} minutes")
