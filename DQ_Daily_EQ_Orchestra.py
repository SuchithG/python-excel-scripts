import os
import pandas as pd
import time
from datetime import datetime, timedelta

start_time = time.time()

# Specify the paths to the folders containing the Excel files
folder_path = "path1"
output_folder_path = "output_path"

def get_previous_working_day():
    today = datetime.now().date()
    if today.weekday() == 0:  # Monday
        return today - timedelta(days=3)
    elif today.weekday() == 6:  # Sunday
        return today - timedelta(days=2)
    else:  # Tuesday to Saturday
        return today - timedelta(days=1)

previous_working_day_str = get_previous_working_day()
print(f"Filtering data for the date: {previous_working_day_str}")

# Create output folder if not exists
if not os.path.exists(output_folder_path):
    os.makedirs(output_folder_path)
    print(f"Created output folder at: {output_folder_path}")

dfs = []

if os.path.exists(folder_path):
    print(f"Processing files in folder: {folder_path}")
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx") and not filename.startswith("~$"):
            file_path = os.path.join(folder_path, filename)
            try:
                df = pd.read_excel(file_path)
                df['Date'] = pd.to_datetime(df['Date'])
                df = df[df['Date'].dt.date == previous_working_day_str]

                if not df.empty:
                    dfs.append(df)
                    print(f"Added data from file: {filename}")
                else:
                    print(f"File {filename} does not contain data for {previous_working_day_str}. Skipping...")
            except Exception as e:
                print(f"Could not read file {filename}. Error: {e}")
else:
    print(f"Folder path '{folder_path}' does not exist. Skipping...")

# Concatenate all data frames
if dfs:
    result_df = pd.concat(dfs, ignore_index=True)
    
    # Remove columns that start with 'Unnamed:'
    result_df = result_df.loc[:, ~result_df.columns.str.startswith('Unnamed:')]

    # Drop rows with empty values in 'Formula', 'Resource Name', 'Date', 'Month'
    result_df.dropna(subset=['Formula', 'Resource Name', 'Date', 'Month'], inplace=True)
    
    # Fill empty cells with zero and convert to numeric
    numeric_columns = ['Setup', 'Amend', 'Review', 'Closure', '4 eye Count', 'Error Count']
    for col in numeric_columns:
        result_df[col] = result_df[col].fillna(0)
        result_df[col] = pd.to_numeric(result_df[col], errors='coerce').astype(int)

    # Save the concatenated data frame to a new Excel file
    output_file_path = os.path.join(output_folder_path, "concatenated_data.xlsx")
    result_df.to_excel(output_file_path, index=False)
    print(f"Data concatenated successfully. Output file saved at: {output_file_path}")
else:
    print("No valid data found. Concatenation skipped.")

end_time = time.time()
execution_time_in_minutes = (end_time - start_time) / 60
print(f"Script execution time: {execution_time_in_minutes:.2f} minutes")
