import pandas as pd
from datetime import datetime
import os

# Get the previous month and year
def get_previous_month_year_str():
    current_date = datetime.now()
    previous_month_date = (current_date.replace(day=1) - pd.Timedelta(days=1))
    return previous_month_date.strftime('%b_%Y')

previous_month_year_str = get_previous_month_year_str()
print(previous_month_year_str)

# Paths to the input files
input_file_path = r'G:\girisuc\RDS Dashboard\RDS monthly tickets inputs\RDS_MONTHLY_REPORT_TEMPLAT_{}.xlsx'.format(previous_month_year_str)
sla_file_path = r'G:\girisuc\RDS Dashboard\RDS SLA.xlsx'

# Paths to the output file
output_file_path = f'G:\\girisuc\\RDS Dashboard\\RDS tickets monthly\\RDS tickets output - {previous_month_year_str}_UAT.xlsx'

# Read the input Excel file
input_df = pd.read_excel(input_file_path)

# Convert date & time columns to the correct format
input_df['Incident Start Date & Time'] = pd.to_datetime(input_df['Incident Start Date & Time']).dt.strftime('%m/%d/%Y %H:%M:%S')
input_df['Incident End Date & Time'] = pd.to_datetime(input_df['Incident End Date & Time']).dt.strftime('%m/%d/%Y %H:%M:%S')

# Read the SLA Excel file
sla_df = pd.read_excel(sla_file_path)

# Convert the 'Incident Duration Excluding GMT Weekends (Seconds)' to hours and round to two decimal places
input_df['Time in HRS'] = round(input_df['Incident Duration Excluding GMT Weekends (Seconds)'] / 3600, 2)

# Rename the 'Incident categorized' column in sla_df to match 'Incident External Reference' in input_df for merging
sla_df.rename(columns={'Incident categorized': 'Incident External Reference'}, inplace=True)

# Merge the input dataframe with the SLA dataframe based on 'Incident External Reference'
merged_df = pd.merge(input_df, sla_df, on='Incident External Reference', how='left')

# Calculate the 'SLA Difference Hrs' as the difference between 'SLA' and 'Time in HRS'
# Rounding the result to two decimal places
merged_df['SLA Difference Hrs'] = round(merged_df['SLA'] - merged_df['Time in HRS'], 2)

# Determine if 'SLA MET/Not Met' based on 'Time in HRS' being less than or equal to 'SLA'
merged_df['SLA MET/Not MET'] = merged_df.apply(
    lambda row: 'SLA Met' if row['Time in HRS'] <= row['SLA'] else 'SLA not Met', axis=1
)

try:
    # Ensure the output directory exists
    output_dir = os.path.dirname(output_file_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Write the processed DataFrame to a new Excel file in the desired location
    merged_df.to_excel(output_file_path, index=False)
    print(f"Processed data has been saved to {output_file_path}")
except Exception as e:
    print(f"An error occurred: {e}")
