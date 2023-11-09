import os
import pandas as pd
import json
import time

start_time = time.time()

def concatenate_and_split_excel_files(input_folder_path, output_folder_path, output_file_name, max_rows_per_sheet=1000000): # Using a number slightly less than the Excel maximum
    # Get a list of all the Excel files in the input folder
    excel_files = [f for f in os.listdir(input_folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

    if not excel_files:
        print("No Excel files found in the specified input folder.")
        return
    
    # Initialize an empty list to hold dataframes
    df_list = []

    # Loop through the list of Excel files and read each one into a pandas DataFrame
    print("Reading Excel files:")
    for excel_file in excel_files:
        print(f"  Reading {excel_file}...")
        df = pd.read_excel(os.path.join(input_folder_path, excel_file))
        df_list.append(df)

    # Concatenate all the dataframes
    print("Concatenating dataframes...")
    concatenated_df = pd.concat(df_list)
    print(f"Total rows in concatenated dataframe: {len(concatenated_df)}")

    # Add new columns with empty values
    concatenated_df["Matching apps"] = ""
    concatenated_df["BOT Scope"] = ""
    concatenated_df["In Accuracy"] = ""
    print("Added new columns: 'Matching apps', 'BOT Scope', and 'In Accuracy'")

    # Create the output folder if it doesn't exist
    os.makedirs(output_folder_path, exist_ok=True)
    print(f"Ensured output folder exists: {output_folder_path}")

    # Save the concatenated dataframe to a new Excel file in the output folder, splitting across sheets
    output_file_path = os.path.join(output_folder_path, output_file_name)
    print(f"Saving concatenated data to {output_file_path}...")
    
    # Splitting the dataframe across multiple sheets
    num_sheets = len(concatenated_df) // max_rows_per_sheet + 1
    print(f"Data will be split across {num_sheets} sheets due to row limits.")
    with pd.ExcelWriter(output_file_path) as writer:
        for i in range(num_sheets):
            print(f"Saving data to Sheet_{i+1}...")
            concatenated_df.iloc[i*max_rows_per_sheet: (i+1)*max_rows_per_sheet].to_excel(writer, sheet_name=f'Sheet_{i+1}', index=False)
    
    print("Concatenation and saving completed successfully.")

# Define the input and output folder paths and output file name
input_folder_path = 'path/to/your/input/folder'
output_folder_path = 'path/to/your/output/folder'
output_file_name = 'output.xlsx'

# Call the function to concatenate all Excel files and save the output
concatenate_and_split_excel_files(input_folder_path, output_folder_path, output_file_name)

end_time = time.time()

execution_time_in_minutes = (end_time - start_time) / 60
print("Script execution in {execution_time_in_minutes:.2f} minutes")