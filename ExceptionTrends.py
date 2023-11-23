import pandas as pd

# Replace 'path_to_your_excel_file.xlsx' with the actual path to your Excel file
file_path = 'path_to_your_excel_file.xlsx'

# Load the data from the Excel file
df = pd.read_excel(file_path, sheet_name='Q1 Deal')

# Ensure the OPEN_DT column is in datetime format
df['OPEN_DT'] = pd.to_datetime(df['OPEN_DT'], errors='coerce')

# Use OPEN_DT for the Month/Year value for all records
df['Month/Year'] = df['OPEN_DT'].dt.strftime('%b-%y')

# Group by MSG_TYP, PRIORITY, STATUS, and Month/Year
grouped = df.groupby(['MSG_TYP', 'PRIORITY', 'STATUS', 'Month/Year'])

# Count the number of occurrences in each group
output_df = grouped.size().reset_index(name='Count')

# Create Status_Count and Volume columns (assuming they are the same)
output_df['Status_Count'] = output_df['Count']
output_df['Volume'] = output_df['Count']

# Group by 'Month/Year' and 'PRIORITY' and count distinct 'ATTR_NME'
grouped_by_month_priority = df.groupby(['Month/Year', 'PRIORITY'])['ATTR_NME'].nunique().reset_index(name='Priority_Count')

# Now to replicate the summarization from Alteryx for distinct counts of ATTR_NME
# Group by MSG_TYP and count distinct ATTR_NME
distinct_attr = df.groupby('MSG_TYP')['ATTR_NME'].nunique().reset_index(name='CountDistinct_ATTR_NME')

# Print the grouped DataFrame with counts for each group
print("Grouped DataFrame by MSG_TYP, PRIORITY, STATUS, and Month/Year:")
print(output_df)

# Print the grouped DataFrame by Month/Year and PRIORITY with distinct counts of ATTR_NME
print("\nGrouped DataFrame by Month/Year and PRIORITY with distinct counts of ATTR_NME:")
print(grouped_by_month_priority)

# Print the distinct count of ATTR_NME per MSG_TYP
print("\nDistinct Count of ATTR_NME per MSG_TYP:")
print(distinct_attr)
