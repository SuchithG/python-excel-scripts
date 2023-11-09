import pandas as pd
from datetime import datetime, timedelta

# Define the base paths for the files
base_path_exceptions = '/path/to/exceptions/folder/'
base_path_dictionaries = '/path/to/dictionaries/folder/'

# Get the previous month's file name based on the current date
previous_month = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%b %y")
dq_exception_file_name = f"DQ_CDE_SOI_{previous_month}.xlsx"

# Construct the full paths for the files
dq_exception_file_path = f"{base_path_exceptions}{dq_exception_file_name}"
asset_class_dict_path = f"{base_path_dictionaries}DQ CDE Dictionary.xlsx"

# Load the 'DQ Exception' sheet and the 'Asset Class Dictionary' sheet
dq_exception_df = pd.read_excel(dq_exception_file_path, sheet_name='DQ Exception')
asset_class_dict_df = pd.read_excel(asset_class_dict_path, sheet_name='Asset Class Dictionary')

# Perform an inner join on the 'MSG_TYP' column
merged_df = pd.merge(dq_exception_df, asset_class_dict_df, on='MSG_TYP')

# Load the 'Concept_Updated' sheet from the "DQ CDE Dictionary.xlsx" file
concept_updated_df = pd.read_excel(asset_class_dict_path, sheet_name='Concept_Updated')

# Perform the second inner join on 'NOTFCN_ID' and 'Asset Class'
final_merged_df = merged_df.merge(concept_updated_df, on=['NOTFCN_ID', 'Asset Class'])

# Select only the specified columns
final_columns = ["COUNT(*)", "MSG_TYP", "NOTFCN_ID", "Asset Class", "Concept"]
final_df = final_merged_df[final_columns]

# Display the first few rows of the resulting DataFrame
print(merged_df.head())