import os
import pandas as pd
import time
from datetime import datetime, timedelta

start_time = time.time()

folder_path = "path1"

output_folder_path = "output_path"

def parse_date(date_input):
    """
    Custom function to handle various date formats and time-only strings.
    """
    # Check for time-only strings and handle them
    if isinstance(date_input, str) and (date_input.strip() in ["00:00:00", "12:00:00 AM", "12:00 AM"] or ':' in date_input):
        # Return None or a default date as needed
        return None  

    # Handle datetime objects and strings with date information
    if isinstance(date_input, datetime):
        return date_input.strftime("%m/%d/%Y")
    elif isinstance(date_input, str):
        for fmt in ("%m/%d/%Y", "%d-%b-%y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(date_input, fmt).strftime("%m/%d/%Y")
            except ValueError:
                continue

    return None

# Function to check if data exists for a given date in a specific file
def data_exists_for_date(file_path, date):
    try:
        df = pd.read_excel(file_path)
        df['Date'] = pd.to_datetime(df['Date'])
        return not df[df['Date'] == date].empty
    except Exception as e:
        print(f"Error checking data in file {file_path}: {e}")
        return False
    
# Function to get the previous working day
def get_previous_working_day(folder_path):
    today = datetime.now()
    if today.weekday() == 0:  # Monday
        last_saturday = today - timedelta(days=2)
        # Check if there is data for last Saturday
        for filename in os.listdir(folder_path):
            if filename.endswith(".xlsm.xlsx") and not filename.startswith("~$"):
                file_path = os.path.join(folder_path, filename)
                if data_exists_for_date(file_path, last_saturday):
                    return last_saturday
        return today - timedelta(days=3)  # Return last Friday
    else:
        return today - timedelta(days=1)

previous_working_day = get_previous_working_day(folder_path)
print(f"Filtering data for the month: {previous_working_day}")

if not os.path.exists(output_folder_path):
    os.makedirs(output_folder_path)
    print(f"Created output 
          folder at: {output_folder_path}")

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
                df = pd.read_excel(file_path, dtype={'Date': str}, converters={'Date': parse_date})
                
                # Check if necessary columns are in the file
                if set(["Formula", "Resource name", "Date", "Month"]).issubset(df.columns):

                    # Covert 'Date' and 'Actual Date of upload' columns
                    df['Date'] = df['Date'].apply(lambda x: parse_date(x) if isinstance(x, str) else x)
                    df['Actual Date of upload'] = df['Actual Date of upload'].dt.strftime("%Y-%m-%d")
                    
                    # Keep only rows where necessary columns are not null
                    df = df[df[["Formula", "Resource name", "Date", "Month"]].notnull().all(axis=1) & (df["Month"] == previous_working_day)]
                    
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
