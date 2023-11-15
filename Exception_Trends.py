import pandas as pd
import datetime
import os

# Function to get the previous month's file name for the input file
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

# Function to extract and combine MSG_TYP values from Q1 Deal and Q1 Tranche
def combine_msg_typ_values(q1_deal, q1_tranche):
    msg_typ_deal = q1_deal['MSG_TYP']
    msg_typ_tranche = q1_tranche['MSG_TYP']
    combined_msg_typ = pd.concat([msg_typ_deal, msg_typ_tranche]).drop_duplicates().reset_index(drop=True)
    return combined_msg_typ

# Function to get the output file name based on the previous month
def get_output_file_name():
    today = datetime.date.today()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - datetime.timedelta(days=1)
    file_date = last_day_of_previous_month.strftime('%b-%y').upper()  # Format as Mon-YY
    return f"ExceptionTrends_{file_date}_script_output.xlsx"

# Function to write data to a new Excel file
def write_to_excel(df, folder_path, sheet_name):
    file_name = get_output_file_name()
    file_path = os.path.join(folder_path, file_name)

    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Data written successfully to sheet '{sheet_name}' in '{file_path}'.")
    except Exception as e:
        print(f"Error writing to Excel: {e}")

# Main script execution
input_folder_path = 'path_to_input_folder'  # Replace with actual path
output_folder_path = 'path_to_output_folder'  # Replace with actual path

input_file_name = get_previous_month_file_name()
input_file_path = os.path.join(input_folder_path, input_file_name)

q1_deal = read_sheet(input_file_path, "Q1 Deal")
q1_tranche = read_sheet(input_file_path, "Q1 Tranche")

if q1_deal is not None and q1_tranche is not None:
    combined_msg_typ = combine_msg_typ_values(q1_deal, q1_tranche)

    output_data = {
        'MSG_TYP': combined_msg_typ,
        # Other columns can be added here with initialization
    }
    output_df = pd.DataFrame(output_data)

    write_to_excel(output_df, output_folder_path, 'Processed Data')
else:
    print("Required input sheets are missing.")
