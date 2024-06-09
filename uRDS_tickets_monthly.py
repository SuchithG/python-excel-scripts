import pandas as pd
from datetime import datetime
import os

# Get the previous month and year as a string
def get_previous_month_year_str():
    current_date = datetime.now()
    previous_month_date = (current_date.replace(day=1) - pd.Timedelta(days=1))
    return previous_month_date.strftime('%b_%Y')

previous_month_year_str = get_previous_month_year_str()
print(previous_month_year_str)

# Define the paths to the input and output files
input_file_path = f'G:\\girisuc\\RDS Dashboard\\RDS monthly tickets inputs\\RDS_MONTHLY_REPORT_TEMPLAT_{previous_month_year_str}.xlsx'
sla_file_path = 'G:\\girisuc\\RDS Dashboard\\RDS SLA.xlsx'
output_file_path = f'G:\\girisuc\\RDS Dashboard\\RDS tickets monthly\\RDS tickets output - {previous_month_year_str}_UAT.xlsx'

# Function to ensure the output directory exists
def ensure_directory_exists(file_path):
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)

# Ensure the output directory exists before proceeding
ensure_directory_exists(output_file_path)

# Load the input DataFrame from the Excel file
input_df = pd.read_excel(input_file_path)

# Convert date & time columns to the correct format
input_df['Incident Start Date & Time'] = pd.to_datetime(input_df['Incident Start Date & Time']).dt.strftime('%m/%d/%Y %H:%M:%S')
input_df['Incident End Date & Time'] = pd.to_datetime(input_df['Incident End Date & Time']).dt.strftime('%m/%d/%Y %H:%M:%S')

# Load the SLA DataFrame from the Excel file
sla_df = pd.read_excel(sla_file_path)

# Process the data as required for the analysis
input_df['Time in HRS'] = round(input_df['Incident Duration Excluding GMT Weekends (Seconds)'] / 3600, 2)
sla_df.rename(columns={'Incident categorized': 'Incident External Reference'}, inplace=True)
merged_df = pd.merge(input_df, sla_df, on='Incident External Reference', how='left')
merged_df['SLA Difference Hrs'] = round(merged_df['SLA'] - merged_df['Time in HRS'], 2)
merged_df['SLA MET/Not MET'] = merged_df.apply(lambda row: 'SLA Met' if row['Time in HRS'] <= row['SLA'] else 'SLA not Met', axis=1)

# Write the processed DataFrame to the Excel file
try:
    merged_df.to_excel(output_file_path, index=False)
    print(f"Processed data has been saved to {output_file_path}")
except Exception as e:
    print(f"An error occurred while saving the file: {e}")
