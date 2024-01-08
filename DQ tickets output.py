import pandas as pd

# Paths to the input files
input_file_path = r'C:\Users\Suchith G\Documents\Test Docs\IRDS_MONTHY_REPORT_TEMPLET.xlsx'
sla_file_path = r'C:\Users\Suchith G\Documents\Test Docs\DQ SLA.xlsx'

# Paths to the output file
output_file_path = r'C:\Users\Suchith G\Documents\Test Docs\DQ Output\DQ tickets output - AugSepOct_UAT.xlsx'

# Read the input Excel file
input_df = pd.read_excel(input_file_path)

# Read the DQ SLA Excel file
sla_df = pd.read_excel(sla_file_path)

# Convert the "Incident Duration Excluding GMT Weekends (Seconds)" to hours and create a new column "Time in HRS"
input_df['Time in HRS'] = input_df['Incident Duration Excluding GMT Weekends (Seconds)'] / 3600

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
