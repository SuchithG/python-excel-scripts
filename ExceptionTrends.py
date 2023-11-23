import pandas as pd

# Replace 'path_to_your_excel_file.xlsx' with the actual path to your Excel file
file_path = 'path_to_your_excel_file.xlsx'

# Load the data from the Excel file
df = pd.read_excel(file_path, sheet_name='Q1 Deal')

# Ensure the OPEN_DT column is in datetime format
df['OPEN_DT'] = pd.to_datetime(df['OPEN_DT'], errors='coerce')

# Use OPEN_DT for the Month/Year value for all records
df['Month/Year'] = df['OPEN_DT'].dt.strftime('%b-%y')

# Group by MSG_TYP, PRIORITY, STATUS, and Month/Year, and count occurrences
grouped = df.groupby(['MSG_TYP', 'PRIORITY', 'STATUS', 'Month/Year'])

# Count the number of occurrences in each group
output_df = grouped.size().reset_index(name='Volume')  # Using 'Volume' here, can be duplicated for 'Status_Count' if needed

# Create Status_Count column, assuming it is the same as Volume
output_df['Status_Count'] = output_df['Volume']

# Now, count distinct ATTR_NME for each MSG_TYP
distinct_attr_count = df.groupby('MSG_TYP')['ATTR_NME'].nunique().reset_index(name='Attribute Count')

# Merge the distinct attribute count with the main DataFrame
# Assuming 'MSG_TYP' is unique in the distinct_attr_count DataFrame
# If not, you might need to adjust the merging strategy
combined_df = pd.merge(output_df, distinct_attr_count, on='MSG_TYP', how='left')

# Print the final DataFrame that matches the Alteryx output
print("Combined DataFrame:")
print(combined_df)
