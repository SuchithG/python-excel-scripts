import pandas as pd
import datetime
import os

def get_previous_month_file_name():
    # Current date
    today = datetime.date.today()

    # First day of the current month
    first_day_of_current_month = today.replace(day=1)

    # Last day of the previous month
    last_day_of_previous_month = first_day_of_current_month - datetime.timedelta(days=1)

    # Format the date to match your file naming convention
    file_date = last_day_of_previous_month.strftime('%b_%Y').upper()

    return f"uRDSandFDW_{file_date}.xlsx"

def read_excel_file(folder_path):
    file_name = get_previous_month_file_name()
    file_path = os.path.join(folder_path, file_name)

    if os.path.exists(file_path):
        # Read the Excel file
        df = pd.read_excel(file_path)
        print(f"File read successfully: {file_path}")
        return df
    else:
        print(f"File not found: {file_path}")
        return None

# Example usage
folder_path = 'path_to_your_folder'
data = read_excel_file(folder_path)

if data is not None:
    # Process your data here
    pass
