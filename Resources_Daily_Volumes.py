import os
import pandas as pd
import time
from datetime import datetime, timedelta

start_time = time.time()

# Specify the paths to the folders containing the excel files
folder_path = "path1"

# Define the output folder path
output_folder_path = "output_path"

# Function to get the previous working day in the required format
def get_previous_working_day():
    today = datetime.now().date() # Use date only, without time
    # subtract one day for Tuesday through Friday: subtract three days for Monday
    if today.weekday() == 0: # If today is Monday
        return today - timedelta(days=3)
    elif today.weekday() == 6: # If today is Sunday
        return today - timedelta(days=2)
    else: # Tuesday through Saturday
        return today - timedelta(days=1)

previous_working_day_str = get_previous_working_day()
print(f"Filtering data for the date: {previous_working_day_str}")

# Create output folder if not exists 
if not os.path.exists(output_folder_path):
    os.makedirs(output_folder_path)
    print(f"Created output folder at: {output_folder_path}")

# List to store data frames
dfs = []

# Data types specification for consistent column data types across all files
data_types = {
    'Formula': str,
    'Resource name': str,
    # Specify other columns and their desired data types, especially those expected to be numeric
    'Count': int, 'Setup': int, 'Amend': int, 'Review': int, 'Closure': int, '4 eye Count': int, 'Error Count': int
    # Uncomment and adjust above line as per your actual data columns and types
}

# Check if folder path exists
if not os.path.exists(folder_path):
    print(f"Folder path '{folder_path}' does not exist. Skipping...")
else:
    print(f"Processing files in folder: {folder_path}")

    # Loop through each file in the folder
    for filename in os.listdir(folder_path):
        
        # Check if the file ends with .xlsm.xlsx and is not a temporary file
        if filename.endswith(".xlsx") and not filename.startswith("~$"):
            
            # Define the full path to the file
            file_path = os.path.join(folder_path, filename)
            
            # Try reading the file with header
            try:
                df = pd.read_excel(file_path)
                
                # Ensure 'Date' column is datetime for proper comparison
                df['Date'] = pd.to_datetime(df['Date'])

                # Keep only rows where 'Date' matches the previous working day
                df = df[df['Date'].dt.date == previous_working_day_str]

                # Ensure all required numeric columns are present and correctly formatted
                numeric_columns = ['Count', 'Setup', 'Amend', 'Review', 'Closure', '4 eye Count', 'Error Count']  # Adjust as needed
                for col in numeric_columns:
                    if col not in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)  # Ensure numeric and clean0
                    else:
                        df[col] = 0

                # Check if necessary columns are in the file
                if set(["Formula", "Resource name", "Date", "Month"]).issubset(df.columns):
                    
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

    # Remove columns that start with "Unnamed:"
    result_df = result_df.loc[:, ~result_df.columns.str.startswith('Unnamed:')] # Remove columns that start with "Unnamed:"

    # Save the concatenated data frame to a new Excel file
    output_file_path = os.path.join(output_folder_path, "concatenated_data.xlsx")
    result_df.to_excel(output_file_path, index=False)
    print(f"Data concatenated successfully. Output file saved at: {output_file_path}")
else:
    print("No valid data found. Concatenation skipped.")

end_time = time.time()

execution_time_in_minutes = (end_time - start_time) / 60
print(f"Script execution in {execution_time_in_minutes:.2f} minutes")
