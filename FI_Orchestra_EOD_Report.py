import pandas as pd
import os
from datetime import datetime

# Get the current year and month
current_year = datetime.now().year
current_month = datetime.now().strftime('%B %Y')

# Construct the input file path dynamically
base_dir = r'G:\{}\FI Exception - {}\SOD and EOD Report'.format(current_year, current_year)
input_file_name = '{} {} SOD and EOD Report Updated Version.xlsx'.format(current_month, current_year)
input_file_path = os.path.join(base_dir, input_file_name)

# Read the Excel file
excel_data = pd.ExcelFile(input_file_path)

# Read data from each sheet
df_leave_tracker = pd.read_excel(input_file_path, sheet_name='Leave Tracker')
df_sales = pd.read_excel(input_file_path, sheet_name='Sales Data')
df_inventory = pd.read_excel(input_file_path, sheet_name='Inventory Data')
df_employee = pd.read_excel(input_file_path, sheet_name='Employee Data')

# Process the Leave Tracker data to include only "Name", "Shift", and "Attendance"
df_leave_tracker_processed = df_leave_tracker[['Name', 'Shift', 'Attendance']]

# Define the file path for the output EOD report
output_file_path = os.path.join(base_dir, 'EOD Report.xlsx')

# Create a new Excel writer object for the EOD report
with pd.ExcelWriter(output_file_path) as writer:
    # Write the processed Leave Tracker data to a sheet
    df_leave_tracker_processed.to_excel(writer, sheet_name='Leave Tracker', index=False)
    # Write each of the other DataFrames to separate sheets
    df_sales.to_excel(writer, sheet_name='Sales Data', index=False)
    df_inventory.to_excel(writer, sheet_name='Inventory Data', index=False)
    df_employee.to_excel(writer, sheet_name='Employee Data', index=False)

print("EOD report has been generated and saved to:", output_file_path)
