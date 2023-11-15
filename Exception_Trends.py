import pandas as pd
import datetime
import os

# Function to get the previous month's file name
def get_previous_month_file_name():
    today = datetime.date.today()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - datetime.timedelta(days=1)
    file_date = last_day_of_previous_month.strftime('%b_%Y').upper()
    return f"uRDSandFDW_{file_date}.xlsx"

# Function to read a sheet from the Excel file
def read_sheet(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"Sheet '{sheet_name}' read successfully.")
        return df
    except Exception as e:
        print(f"Error reading '{sheet_name}': {e}")
        return None

# Process the data here (Placeholder functions)
def process_q1_q2_deal(df):
    # Add your processing steps here
    pass

def process_q1_q2_tranche(df):
    # Add your processing steps here
    pass

# Function to get the current month's output file name
def get_output_file_name():
    today = datetime.date.today()
    file_date = today.strftime('%b-%y').upper()
    return f"ExceptionTrends_{file_date}_script_output.xlsx"

# Function to write data to an Excel file
def write_to_excel(df, folder_path, sheet_name):
    file_name = get_output_file_name()
    file_path = os.path.join(folder_path, file_name)

    try:
        with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Data written successfully to sheet '{sheet_name}' in '{file_path}'.")
    except Exception as e:
        print(f"Error writing to Excel: {e}")

# Main script execution
input_folder_path = 'path_to_input_folder'
output_folder_path = 'path_to_output_folder'

input_file_name = get_previous_month_file_name()
input_file_path = os.path.join(input_folder_path, input_file_name)

# Reading each sheet
q1_deal = read_sheet(input_file_path, "Q1 Deal")
q2_deal = read_sheet(input_file_path, "Q2 Deal")
q1_tranche = read_sheet(input_file_path, "Q1 Tranche")
q2_tranche = read_sheet(input_file_path, "Q2 Tranche")

# Process each DataFrame as needed
# Replace these with your actual processing functions
processed_data = {
    'Q1 Deal': process_q1_q2_deal(q1_deal) if q1_deal is not None else None,
    'Q2 Deal': process_q1_q2_deal(q2_deal) if q2_deal is not None else None,
    'Q1 Tranche': process_q1_q2_tranche(q1_tranche) if q1_tranche is not None else None,
    'Q2 Tranche': process_q1_q2_tranche(q2_tranche) if q2_tranche is not None else None
}

# Writing processed data to the output Excel file
for sheet_name, df in processed_data.items():
    if df is not None:
        write_to_excel(df, output_folder_path, sheet_name)
