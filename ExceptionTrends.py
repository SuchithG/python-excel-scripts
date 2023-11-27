import os
import pandas as pd
from datetime import datetime, timedelta

previous_month_date = (datetime.now().replace(day=1) - timedelta(days=1))
previous_month_name = previous_month_date.strftime("%b_%Y").upper()

# Replace 'path_to_your_excel_file.xlsx' with the actual path to your Excel file
file_path = 'path_to_your_excel_file.xlsx'

# Load the data from the Excel file's "Q1 Deal" sheet
df_q1 = pd.read_excel(file_path, sheet_name='Q1 Deal')

# Load the "COUNT(*)" column from the "Q2 Deal" sheet
df_q2 = pd.read_excel(file_path, sheet_name='Q2 Deal', usecols=['COUNT(*)'])

# Get the first value from the "COUNT(*)" column of "Q2 Deal" sheet (assuming it's the same for all records)
count_value = df_q2['COUNT(*)'].iloc[0]

# Ensure the OPEN_DT column in "Q1 Deal" is in datetime format
df_q1['OPEN_DT'] = pd.to_datetime(df_q1['OPEN_DT'], errors='coerce')

# Use OPEN_DT for the Month/Year value for all records
df_q1['Month/Year'] = df_q1['OPEN_DT'].dt.strftime('%b-%y')

msg_typ_rename = {
    '3DS': '3D Static',
    'BF': 'Business Finance',
    'BBG': 'Bloomberg'
}

df_q1['MSG_TYP'] = df_q1['MSG_TYP'].map(msg_typ_rename).fillna(df_q1['MSG_TYP'])

# Group by MSG_TYP, PRIORITY, STATUS, and Month/Year, and count occurrences
grouped = df_q1.groupby(['MSG_TYP', 'PRIORITY', 'STATUS', 'Month/Year'])

# Count the number of occurrences in each group
output_df = grouped.size().reset_index(name='Volume')

# Create Status_Count column, assuming it is the same as Volume
output_df['Status_Count'] = output_df['Volume']

# Now, count distinct ATTR_NME for each MSG_TYP
distinct_attr_count = df_q1.groupby('MSG_TYP')['ATTR_NME'].nunique().reset_index(name='Attribute Count')

# Calculate Priority Count which is the count of distinct ATTR_NME for a given PRIORITY across all MSG_TYP
priority_count = df_q1.groupby('PRIORITY')['ATTR_NME'].nunique().reset_index(name='Priority Count')

# Merge the distinct attribute count with the main DataFrame
combined_df = pd.merge(output_df, distinct_attr_count, on='MSG_TYP', how='left')

# Merge the priority count with the main DataFrame
combined_df = pd.merge(combined_df, priority_count, on='PRIORITY', how='left')

# Add column 'Count(*)' with the value from "Q2 Deal" sheet
combined_df['Count(*)'] = count_value

# Determine 'Exception trend' based on the sheet data we are processing
combined_df['Exception trend'] = 'Deals'

banking_book_types = ["Business Finance", "Paragon", "TAS"]
trading_book_types = ["3D Static", "Bloomberg", "Intex", "STS"]

combined_df['Message Type Group'] = combined_df['MSG_TYP'].apply(
    lambda x: "Banking Book" if x in banking_book_types else "Trading Book" if x in trading_book_types else "Other"
)

# Reorder columns to match the provided screenshot, excluding 'Right Exception trend'
final_df = combined_df[['Count(*)', 'Exception trend', 'MSG_TYP', 'Message Type Group', 'Attribute Count', 
                        'PRIORITY', 'Priority Count', 'Volume', 'Month/Year', 'STATUS', 'Status_Count']]

# Load the data from the Excel file's "Q1 Tranche" sheet
df_q1_tranche = pd.read_excel(file_path, sheet_name='Q1 Tranche')

# Load the "COUNT(*)" column from the "Q2 Tranche" sheet
df_q2_tranche = pd.read_excel(file_path, sheet_name='Q2 Tranche', usecols=['COUNT(*)'])

# Get the first value from the "COUNT(*)" column of "Q2 Deal" sheet (assuming it's the same for all records)
count_value_tranche = df_q2_tranche['COUNT(*)'].iloc[0]

# Ensure the OPEN_DT column in "Q1 Deal" is in datetime format
df_q1_tranche['OPEN_DT'] = pd.to_datetime(df_q1_tranche['OPEN_DT'], errors='coerce')

# Use OPEN_DT for the Month/Year value for all records
df_q1_tranche['Month/Year'] = df_q1_tranche['OPEN_DT'].dt.strftime('%b-%y')

msg_typ_rename = {
    '3DS': '3D Static',
    'BF': 'Business Finance',
    'BBG': 'Bloomberg'
}

df_q1_tranche['MSG_TYP'] = df_q1_tranche['MSG_TYP'].map(msg_typ_rename).fillna(df_q1_tranche['MSG_TYP'])

# Group by MSG_TYP, PRIORITY, STATUS, and Month/Year, and count occurrences
grouped_tranche = df_q1_tranche.groupby(['MSG_TYP', 'PRIORITY', 'STATUS', 'Month/Year'])

# Count the number of occurrences in each group
output_df_tranche = grouped_tranche.size().reset_index(name='Volume')

# Create Status_Count column, assuming it is the same as Volume
output_df_tranche['Status_Count'] = output_df_tranche['Volume']

# Now, count distinct ATTR_NME for each MSG_TYP
distinct_attr_count_tranche = df_q1_tranche.groupby('MSG_TYP')['ATTR_NME'].nunique().reset_index(name='Attribute Count')

# Calculate Priority Count which is the count of distinct ATTR_NME for a given PRIORITY across all MSG_TYP
priority_count_tranche = df_q1_tranche.groupby('PRIORITY')['ATTR_NME'].nunique().reset_index(name='Priority Count')

# Merge the distinct attribute count with the main DataFrame
combined_df_tranche = pd.merge(output_df_tranche, distinct_attr_count_tranche, on='MSG_TYP', how='left')

# Merge the priority count with the main DataFrame
combined_df_tranche = pd.merge(combined_df_tranche, priority_count_tranche, on='PRIORITY', how='left')

# Add column 'Count(*)' with the value from "Q2 Tranche" sheet
combined_df_tranche['Count(*)'] = count_value_tranche

# Determine 'Exception trend' based on the sheet data we are processing
combined_df_tranche['Exception trend'] = 'Tranche'

banking_book_types_tranche = ["Business Finance", "Paragon", "TAS"]
trading_book_types_tranche = ["3D Static", "Bloomberg", "Intex", "STS"]

combined_df['Message Type Group'] = combined_df_tranche['MSG_TYP'].apply(
    lambda x: "Banking Book" if x in banking_book_types_tranche else "Trading Book" if x in trading_book_types_tranche else "Other"
)

# Reorder columns to match the provided screenshot, excluding 'Right Exception trend'
final_df_tranche = combined_df_tranche[['Count(*)', 'Exception trend', 'MSG_TYP', 'Message Type Group', 'Attribute Count', 
                        'PRIORITY', 'Priority Count', 'Volume', 'Month/Year', 'STATUS', 'Status_Count']]

final_combined_df = pd.concat([final_df, final_df_tranche], ignore_index=True)

final_combined_df.rename(columns={
    'Count(*)': 'Total'
}, inplace=True)

ordered_columns = [
    'Exception trend', 'MSG_TYP', 'Message Type Group', 'Attribute Count', 
    'PRIORITY', 'Volume', 'Total', 'Month/Year', 'STATUS', 'Priority Count'
]
final_combined_df = final_combined_df[ordered_columns]

# Specify the output folder path and file name
output_folder_path = '/path/to/your/output/folder/'
output_file_name = 'final_output.xlsx'

# Create the full file path
full_file_path = os.path.join(output_folder_path, output_file_name)

# Check if the output folder exists, and if not, create it
if not os.path.exists(output_folder_path):
    os.makedirs(output_folder_path)

# Save the DataFrame to an Excel file
final_combined_df.to_excel(full_file_path, index=False)

print(f"File saved successfully to {full_file_path}")