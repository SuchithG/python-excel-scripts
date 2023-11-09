import os
import pandas as pd
from datetime import datetime, timedelta

# Define the base paths for the files
base_path_exceptions = '/path/to/exceptions/folder/'
base_path_dictionaries = '/path/to/dictionaries/folder/'

# Get the previous month's date
previous_month_date = datetime.now().replace(day=1) - timedelta(days=1)
# Format the previous month's name and year
previous_month_name = previous_month_date.strftime("%B-%Y")

# Load the 'DQ Exception' sheet from the Excel file
dq_exception_file_name = f"DQ_CDE_SOI_{previous_month_date.strftime('%b %y')}.xlsx"

# Construct the full paths for the files
dq_exception_file_path = f"{base_path_exceptions}{dq_exception_file_name}"
asset_class_dict_path = f"{base_path_dictionaries}DQ CDE Dictionary.xlsx"

# Load the 'DQ Exception' sheet and the 'Asset Class Dictionary' sheet
dq_exception_df = pd.read_excel(dq_exception_file_path, sheet_name='DQ Exception')
asset_class_dict_df = pd.read_excel(asset_class_dict_path, sheet_name='Asset Class Dictionary')

# Perform an inner join on the 'MSG_TYP' column
merged_df_1 = pd.merge(dq_exception_df, asset_class_dict_df, on='MSG_TYP')

# Load the 'Concept_Updated' sheet from the Excel file
concept_updated_df = pd.read_excel(asset_class_dict_path, sheet_name='Concept_Updated')

# Perform the second inner join on 'NOTFCN_ID' and 'Asset Class'
final_merged_df = merged_df_1.merge(concept_updated_df, on=['NOTFCN_ID', 'Asset Class'])

# Define valid concepts and asset classes
valid_concepts = ["Accuracy", "Completeness", "Conformity", "Consistency", "Timeliness", "Uniqueness"]
valid_asset_classes = ['Equity', 'LD', 'FI']


# Filter 'final_merged_df' to include only valid concepts and asset classes
final_merged_df = final_merged_df[
    final_merged_df['Concept'].isin(valid_concepts) & 
    final_merged_df['Asset Class'].isin(valid_asset_classes)
]

# Group by 'Asset Class' and 'Concept', and get the sum of 'COUNT(*)'
grouped_df = final_merged_df.groupby(['Asset Class', 'Concept'])['COUNT(*)'].sum().reset_index()

# Rename the 'COUNT(*)' column to 'Exception Numbers'
grouped_df.rename(columns={'COUNT(*)': 'Exception Numbers'}, inplace=True)

# Load the 'DQ SOI' sheet from the same Excel file
dq_soi_df = pd.read_excel(dq_exception_file_path, sheet_name='DQ SOI')

# Define a mapping from 'ASSET_CLASS' values in 'DQ SOI' to 'Asset Class' values in 'grouped_df'
asset_class_mapping = {
    'FixedIncome': 'FI',
    'Equity': 'Equity',
    'ListedDerivative': 'LD'
}

# Update 'Universe Numbers' in 'grouped_df' based on 'ASSET_CLASS' values from 'dq_soi_df'
for soi_asset_class, group_asset_class in asset_class_mapping.items():
    universe_number = dq_soi_df.loc[dq_soi_df['ASSET_CLASS'] == soi_asset_class, 'COUNT(*)'].squeeze()
    # Convert the 'Universe Numbers' to int to avoid floating point representation
    universe_number = int(universe_number) if pd.notna(universe_number) else 0
    grouped_df.loc[grouped_df['Asset Class'] == group_asset_class, 'Universe Numbers'] = universe_number

# Ensure the 'Universe Numbers' column is of type int
grouped_df['Universe Numbers'] = grouped_df['Universe Numbers'].astype(int)

# Add a new column 'Month' at the beginning with the previous month's name
grouped_df.insert(0, 'Month', previous_month_name)

# Add the new column 'Month' at the beginning with the previous month's name
grouped_df.insert(0, 'Month', previous_month_name)

# Construct the output file name based on the previous month
output_file_name = f"DQ_CDE_{previous_month_date.strftime('%b-%y')}_script_output.xlsx"

# Define the output file path
output_file_path = f'/desired/path/to/output/folder/{output_file_name}'

# Save the resulting DataFrame to an Excel file
grouped_df.to_excel(output_file_path, index=False)

print(f"DataFrame has been saved to {output_file_path}")

# Define the source directory where all Excel files are located
source_directory = '/path/to/source/folder/'

# Define the target directory where the concatenated file will be saved
target_directory = '/path/to/target/folder/'

# List all Excel files in the source directory
excel_files = [file for file in os.listdir(source_directory) if file.endswith('.xlsx')]

# Read each Excel file into a DataFrame and store them in a list
df_list = [pd.read_excel(os.path.join(source_directory, file)) for file in excel_files]

# Concatenate all DataFrames into one
concatenated_df = pd.concat(df_list, ignore_index=True)

# Convert 'Month' column to datetime for sorting
concatenated_df['Month'] = pd.to_datetime(concatenated_df['Month'], format='%B-%Y')

# Sort the DataFrame based on the datetime 'Month' column
concatenated_df.sort_values(by='Month', inplace=True)

# Convert the datetime 'Month' column back to the desired string format if needed
concatenated_df['Month'] = concatenated_df['Month'].dt.strftime('%B-%Y')

# Define the output file name for the concatenated DataFrame
output_file_name_concat = "DQ_Metrics_SOI_UAT.xlsx"

# Save the concatenated DataFrame to a new Excel file in the target folder
concatenated_df.to_excel(os.path.join(target_directory, output_file_name_concat), index=False)

print(f"All Excel files have been concatenated and saved to {os.path.join(target_directory, output_file_name_concat)}")