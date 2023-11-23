import pandas as pd

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

# Group by MSG_TYP, PRIORITY, STATUS, and Month/Year, and count occurrences
grouped = df_q1.groupby(['MSG_TYP', 'PRIORITY', 'STATUS', 'Month/Year'])

# Count the number of occurrences in each group
output_df = grouped.size().reset_index(name='Volume')

# Create Status_Count column, assuming it is the same as Volume
output_df['Status_Count'] = output_df['Volume']

# Now, count distinct ATTR_NME for each MSG_TYP
distinct_attr_count = df_q1.groupby('MSG_TYP')['ATTR_NME'].nunique().reset_index(name='Attribute Count')

# Merge the distinct attribute count with the main DataFrame
combined_df = pd.merge(output_df, distinct_attr_count, on='MSG_TYP', how='left')

# Calculate Priority Count which is the count of distinct ATTR_NME for a given PRIORITY across all MSG_TYP
priority_count = df_q1.groupby('PRIORITY')['ATTR_NME'].nunique().reset_index(name='Priority Count')

# Merge the priority count with the main DataFrame
combined_df = pd.merge(combined_df, priority_count, on='PRIORITY', how='left')

# Add column 'Count(*)' with the value from "Q2 Deal" sheet
combined_df['Count(*)'] = count_value

# Determine 'Exception trend' based on the sheet data we are processing
combined_df['Exception trend'] = 'Q1 Deal'

# Reorder columns to match the provided screenshot, excluding 'Right Exception trend'
final_df = combined_df[['Count(*)', 'Exception trend', 'MSG_TYP', 'Attribute Count', 
                        'PRIORITY', 'Priority Count', 'Volume', 'Month/Year', 'STATUS', 'Status Count']]

# Load the data from the Excel file's "Q1 Tranche" sheet
df_q1_tranche = pd.read_excel(file_path, sheet_name='Q1 Tranche')

# Load the "COUNT(*)" column from the "Q2 Tranche" sheet
df_q2_tranche = pd.read_excel(file_path, sheet_name='Q2 Tranche', usecols=['COUNT(*)'])

# Get the first value from the "COUNT(*)" column of "Q2 Deal" sheet (assuming it's the same for all records)
count_value_2 = df_q2_tranche['COUNT(*)'].iloc[0]

# Ensure the OPEN_DT column in "Q1 Deal" is in datetime format
df_q1_tranche['OPEN_DT'] = pd.to_datetime(df_q1_tranche['OPEN_DT'], errors='coerce')

# Use OPEN_DT for the Month/Year value for all records
df_q1_tranche['Month/Year'] = df_q1_tranche['OPEN_DT'].dt.strftime('%b-%y')

# Group by MSG_TYP, PRIORITY, STATUS, and Month/Year, and count occurrences
grouped1 = df_q1_tranche.groupby(['MSG_TYP', 'PRIORITY', 'STATUS', 'Month/Year'])

# Count the number of occurrences in each group
output_df_2 = grouped1.size().reset_index(name='Volume')

# Create Status_Count column, assuming it is the same as Volume
output_df_2['Status_Count'] = output_df_2['Volume']

# Now, count distinct ATTR_NME for each MSG_TYP
distinct_attr_count_1 = df_q1_tranche.groupby('MSG_TYP')['ATTR_NME'].nunique().reset_index(name='Attribute Count')

# Merge the distinct attribute count with the main DataFrame
combined_df_2 = pd.merge(output_df_2, distinct_attr_count_1, on='MSG_TYP', how='left')

# Calculate Priority Count which is the count of distinct ATTR_NME for a given PRIORITY across all MSG_TYP
priority_count_2 = df_q1_tranche.groupby('PRIORITY')['ATTR_NME'].nunique().reset_index(name='Priority Count')

# Merge the priority count with the main DataFrame
combined_df_2 = pd.merge(combined_df_2, priority_count_2, on='PRIORITY', how='left')

# Add column 'Count(*)' with the value from "Q2 Tranche" sheet
combined_df_2['Count(*)'] = count_value_2

# Determine 'Exception trend' based on the sheet data we are processing
combined_df_2['Exception trend'] = 'Q2 Tranche'

# Reorder columns to match the provided screenshot, excluding 'Right Exception trend'
final_df_2 = combined_df_2[['Count(*)', 'Exception trend', 'MSG_TYP', 'Attribute Count', 
                        'PRIORITY', 'Priority Count', 'Volume', 'Month/Year', 'STATUS', 'Status_Count']]

final_combined_df = pd.concat([final_df, final_df_2], ignore_index=True)

# Print the final DataFrame
print("Final Combined DataFrame:")
print(final_combined_df)
