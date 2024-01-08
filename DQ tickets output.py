import pandas as pd
from datetime import datetime

# Function to get the previous month and year in the format "Dec_23"
def get_previous_month_year_str():
    current_date = datetime.now()
    previous_month_date = current_date.replace(day=1) - pd.Timedelta(days=1)
    return previous_month_date.strftime('%b_%y')

# Call the function to get the string for the previous month and year
previous_month_year_str = get_previous_month_year_str()

# Paths to the input files
input_file_path = f'C:\Users\Suchith G\Documents\Test Docs\IRDS_MONTHY_REPORT_TEMPLET_{{{previous_month_year_str}}}.xlsx'
sla_file_path = r'C:\Users\Suchith G\Documents\Test Docs\DQ SLA.xlsx'

# Paths to the output file
output_file_path = f'C:\Users\Suchith G\Documents\Test Docs\DQ Output\DQ tickets output - {previous_month_year_str}_UAT.xlsx'

# Read the input Excel file
input_df = pd.read_excel(input_file_path)

# Read the DQ SLA Excel file
sla_df = pd.read_excel(sla_file_path)

# Convert the "Incident Duration Excluding GMT Weekends (Seconds)" to hours and round to two decimal places
input_df['Time in HRS'] = round(input_df['Incident Duration Excluding GMT Weekends (Seconds)'] / 3600, 2)

# Rename the 'SLA Desc' column in sla_df to match 'Incident External Reference' in input_df for merging
sla_df.rename(columns={'SLA Desc': 'Incident External Reference'}, inplace=True)

# Merge the input dataframe with the SLA dataframe based on "Incident External Reference"
merged_df = pd.merge(input_df, sla_df, on='Incident External Reference', how='left')

# Calculate the "SLA Difference Hrs" as the difference between "SLA" and "Time in HRS"
# Rounding the result to two decimal places
merged_df['SLA Difference Hrs'] = round(merged_df['SLA'] - merged_df['Time in HRS'], 2)

# Determine if "SLA MET/Not MET" based on "Time in HRS" being less than or equal to "SLA"
merged_df['SLA MET/Not MET'] = merged_df.apply(
    lambda row: 'SLA Met' if row['Time in HRS'] <= row['SLA'] else 'SLA Not Met', axis=1
)

# Write the processed DataFrame to a new Excel file in the desired location
merged_df.to_excel(output_file_path, index=False)

print(f"Processed data has been saved to {output_file_path}")
