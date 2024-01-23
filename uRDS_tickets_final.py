import os
import pandas as pd

def concatenate_excel_files(input_folder_path, output_folder_path, output_file_name):
    print("Starting the process of concatenating Excel files...")

    # List all Excel files in the input folder
    excel_files = [file for file in os.listdir(input_folder_path) if file.endswith('.xlsx')]

    if not excel_files:
        print("No Excel files found in the input folder.")
        return

    print(f"Found {len(excel_files)} Excel files to concatenate.")

    # Initialize an empty list to store the dataframes
    dfs = []

    # Read each Excel file and append to the list
    for file in excel_files:
        file_path = os.path.join(input_folder_path, file)
        try:
            df = pd.read_excel(file_path)
            dfs.append(df)
            print(f"Successfully read file: {file}")
        except Exception as e:
            print(f"Error reading {file}: {e}")

    # Check if any dataframes were added
    if not dfs:
        print("No data was read from the files. Exiting.")
        return

    # Concatenate all dataframes
    concatenated_df = pd.concat(dfs, ignore_index=True)
    print("Successfully concatenated all files.")

    # Save the concatenated dataframe to a new Excel file in the output folder
    output_path = os.path.join(output_folder_path, output_file_name)
    try:
        concatenated_df.to_excel(output_path, index=False, sheet_name='Sheet1')
        print(f"Concatenated Excel file saved as '{output_file_name}' in '{output_folder_path}'.")
    except Exception as e:
        print(f"Error saving the concatenated file: {e}")

# Example usage
input_folder = 'path/to/input/folder'  # Replace with your input folder path
output_folder = 'path/to/output/folder'  # Replace with your output folder path
output_file_name = 'dbRIB Tickets-uRDS UAT.xlsx'

concatenate_excel_files(input_folder, output_folder, output_file_name)
