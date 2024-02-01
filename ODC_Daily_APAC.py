import os
import pandas as pd
import time
from datetime import datetime, timedelta

start_time = time.time()

folder_path = "path1"
output_folder_path = "output_path"

def get_previous_working_day():
    today = datetime.now()
    offset = 1 if today.weekday() > 0 else 3  # if today is Monday, offset by 3, otherwise 1
    previous_day = today - timedelta(days=offset)
    return previous_day.strftime("%d-%b-%y")  # Format as 'dd-Mon-yy'

previous_working_day_str = get_previous_working_day()
print(f"Filtering data for the date: {previous_working_day_str}")

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
                df = pd.read_excel(file_path, parse_dates=["Date"], date_parser=lambda x: pd.to_datetime(x, format='%d-%b-%y'))
                
                # Check if necessary columns are in the file
                if set(["Formula", "Resource name", "Date", "Month"]).issubset(df.columns):
                    
                    # Convert 'Date' column to string for comparison
                    df['Date'] = df['Date'].dt.strftime("%d-%b-%y")
                    
                    # Keep only rows where 'Date' matches the previous working day
                    df = df[df['Date'] == previous_working_day_str]
                    
                    if not df.empty:
                        dfs.append(df)
                        print(f"Added data from file: {filename}")
                    else:
                        print(f"File {filename} does not contain data for {previous_working_day_str}. Skipping...")
            
            except Exception as e:
                print(f"Could not read file {filename}. Error: {e}")

# After concatenating all data frames
if dfs:
    result_df = pd.concat(dfs, ignore_index=True)

    # Replace NaN values with 0
    result_df.fillna(0, inplace=True)

    # Save the concatenated data frame to a new Excel file
    output_file_path = os.path.join(output_folder_path, "concatenated_data.xlsx")
    result_df.to_excel(output_file_path, index=False)
    print(f"Data concatenated successfully. Output file saved at: {output_file_path}")
else:
    print("No valid data found. Concatenation skipped.")

end_time = time.time()

execution_time_in_minutes = (end_time - start_time) / 60
print(f"Script execution in {execution_time_in_minutes:.2f} minutes")
