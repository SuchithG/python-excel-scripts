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
merged_df_1 = pd.merge(dq_exception_df, asset_class_dict_df, on='MSG_TYP')
print("Headers after first join:", merged_df_1.columns.tolist())

# Load the 'Concept_Updated' sheet from the Excel file
concept_updated_df = pd.read_excel(asset_class_dict_path, sheet_name='Concept_Updated')

# Perform the second inner join on 'NOTFCN_ID' and 'Asset Class'
final_merged_df = merged_df_1.merge(concept_updated_df, on=['NOTFCN_ID', 'Asset Class'])
print("Headers after second join:", final_merged_df.columns.tolist())

# Group by 'Asset Class' and 'Concept', and get the sum of 'COUNT(*)'
grouped_df = final_merged_df.groupby(['Asset Class', 'Concept'])['COUNT(*)'].sum().reset_index()

# Print the resulting DataFrame
print(grouped_df)