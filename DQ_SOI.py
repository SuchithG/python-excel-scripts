import pandas as pd

# Define the paths to the Excel files
dq_exception_path = 'path/to/DQ_CDE_SOI_23_Oct 23.xlsx'
asset_class_dict_path = 'path/to/DQ CDE Dictionary.xlsx'

# Load the 'DQ Exception' sheet from the first Excel file
dq_exception_df = pd.read_excel(dq_exception_path, sheet_name='DQ Exception')

# Load the 'Asset Class Dictionary' sheet from the second Excel file
asset_class_dict_df = pd.read_excel(asset_class_dict_path, sheet_name='Asset Class Dictionary')

# Perform an inner join on the 'MSG_TYP' column
merged_df = pd.merge(dq_exception_df, asset_class_dict_df, on='MSG_TYP', how='inner')

# Display the first few rows of the resulting DataFrame
print(merged_df.head())