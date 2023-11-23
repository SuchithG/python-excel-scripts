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

# Now to replicate the summarization from Alteryx for distinct counts of ATTR_NME
# Group by MSG_TYP and count distinct ATTR_NME
distinct_attr = df.groupby('MSG_TYP')['ATTR_NME'].nunique().reset_index(name='CountDistinct_ATTR_NME')

# Print the results
print("Grouped DataFrame:")
print(output_df)
print("\nDistinct Count of ATTR_NME per MSG_TYP:")
print(distinct_attr)
