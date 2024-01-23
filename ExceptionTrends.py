import pandas as pd
from itertools import product
from datetime import datetime, timedelta
import glob
import os
import time

# Define previous month details
previous_month_date = (datetime.now().replace(day=1) - timedelta(days=1))
previous_month_name = previous_month_date.strftime("%b_%Y").upper()
uniform_month_year = previous_month_date.strftime("%b-%y") # Uniform Month/Year for all rows

# Load the data from the Excel file
file_path = f"G:/girisuc/DQ CDE's Metrics and FDW Securitization/Exception_Trends_Monthly_Inputs/uRDSandFDW_{previous_month_name}.xlsx"  # Replace with the actual path
df_q1 = pd.read_excel(file_path, sheet_name='Q1 Deal')
df_q2 = pd.read_excel(file_path, sheet_name='Q2 Deal', usecols=['COUNT(*)'])

# Get the first value from the "COUNT(*)" column
total_value = df_q2['COUNT(*)'].iloc[0]

# Process the data
df_q1['OPEN_DT'] = pd.to_datetime(df_q1['OPEN_DT'], errors='coerce')

msg_typ_rename = {
    '3DS': '3D Static',
    'BF': 'Business Finance',
    'BBG': 'Bloomberg',
    # Add other mappings as needed
}
df_q1['MSG_TYP'] = df_q1['MSG_TYP'].map(msg_typ_rename).fillna(df_q1['MSG_TYP'])

# Group and count
grouped = df_q1.groupby(['MSG_TYP', 'PRIORITY', 'STATUS']).size().reset_index(name='Volume')

# Calculate the Attribute Count for each MSG_TYP
attribute_count_per_msg_type = df_q1.groupby('MSG_TYP')['ATTR_NME'].nunique().reset_index(name='Attribute Count')

# Calculate the Priority Count for each PRIORITY
priority_count = df_q1.groupby('PRIORITY')['ATTR_NME'].nunique().reset_index(name='Priority Count')

# Define MSG_TYP and PRIORITY categories
all_msg_types = ["3D Static", "Bloomberg", "Business Finance", "Intex", "Paragon", "STS", "TAS"]
all_priorities = ['P1', 'P2', 'P3']
status_types = ['OPEN', 'CLOSED']

# Create a dataframe with all combinations of MSG_TYP, PRIORITY, and STATUS
all_combinations = pd.DataFrame(product(all_msg_types, all_priorities, status_types), columns=['MSG_TYP', 'PRIORITY', 'STATUS'])

# Merge with the existing dataframe to ensure all combinations are present
full_combined_df = pd.merge(all_combinations, grouped, on=['MSG_TYP', 'PRIORITY', 'STATUS'], how='left')

# Fill missing values in Volume with 0
full_combined_df['Volume'] = full_combined_df['Volume'].fillna(0)

# Merge the calculated Attribute Count and Priority Count back into the full_combined_df
full_combined_df = pd.merge(full_combined_df, attribute_count_per_msg_type, on='MSG_TYP', how='left')
full_combined_df = pd.merge(full_combined_df, priority_count, on='PRIORITY', how='left')

# Fill missing values in Attribute Count with 0
full_combined_df['Attribute Count'] = full_combined_df['Attribute Count'].fillna(0)

# Set Month/Year to be the same for all MSG_TYP values
full_combined_df['Month/Year'] = previous_month_name

# Define banking and trading book types
banking_book_types = ["Business Finance", "Paragon", "TAS"]
trading_book_types = ["3D Static", "Bloomberg", "Intex", "STS"]

full_combined_df['Message Type Group'] = full_combined_df['MSG_TYP'].apply(
    lambda x: "Banking Book" if x in banking_book_types else "Trading Book" if x in trading_book_types else "Other"
)

# Add 'Exception trend' and 'Total' columns
full_combined_df['Exception trend'] = 'Deals'
full_combined_df['Total'] = total_value

# Reorder and select relevant columns as per the specified order
final_df = full_combined_df[['Exception trend', 'MSG_TYP', 'Message Type Group', 'Attribute Count', 
                             'PRIORITY', 'Volume', 'Total', 'Month/Year', 'STATUS', 'Priority Count']]

df_q1_tranche = pd.read_excel(file_path, sheet_name='Q1 Tranche')
df_q2_tranche = pd.read_excel(file_path, sheet_name='Q2 Tranche', usecols=['COUNT(*)'])

# Get the first value from the "COUNT(*)" column
total_value_tranche = df_q2_tranche['COUNT(*)'].iloc[0]

# Process the data
df_q1['OPEN_DT'] = pd.to_datetime(df_q1['OPEN_DT'], errors='coerce')

msg_typ_rename_tranche = {
    '3DS': '3D Static',
    'BF': 'Business Finance',
    'BBG': 'Bloomberg',
    # Add other mappings as needed
}
df_q1_tranche['MSG_TYP'] = df_q1_tranche['MSG_TYP'].map(msg_typ_rename_tranche).fillna(df_q1_tranche['MSG_TYP'])

# Group and count
grouped_tranche = df_q1_tranche.groupby(['MSG_TYP', 'PRIORITY', 'STATUS']).size().reset_index(name='Volume')

# Calculate the Attribute Count for each MSG_TYP
attribute_count_per_msg_type_tranche = df_q1_tranche.groupby('MSG_TYP')['ATTR_NME'].nunique().reset_index(name='Attribute Count')

# Calculate the Priority Count for each PRIORITY
priority_count_tranche = df_q1_tranche.groupby('PRIORITY')['ATTR_NME'].nunique().reset_index(name='Priority Count')

# Define MSG_TYP and PRIORITY categories
all_msg_types_tranche = ["3D Static", "Bloomberg", "Business Finance", "Intex", "Paragon", "STS", "TAS"]
all_priorities_tranche = ['P1', 'P2', 'P3']
status_types_tranche = ['OPEN', 'CLOSED']

# Create a dataframe with all combinations of MSG_TYP, PRIORITY, and STATUS
all_combinations_tranche = pd.DataFrame(product(all_msg_types_tranche, all_priorities_tranche, status_types_tranche), columns=['MSG_TYP', 'PRIORITY', 'STATUS'])

# Merge with the existing dataframe to ensure all combinations are present
full_combined_df_tranche = pd.merge(all_combinations_tranche, grouped_tranche, on=['MSG_TYP', 'PRIORITY', 'STATUS'], how='left')

# Fill missing values in Volume with 0
full_combined_df_tranche['Volume'] = full_combined_df_tranche['Volume'].fillna(0)

# Merge the calculated Attribute Count and Priority Count back into the full_combined_df
full_combined_df_tranche = pd.merge(full_combined_df_tranche, attribute_count_per_msg_type_tranche, on='MSG_TYP', how='left')
full_combined_df_tranche = pd.merge(full_combined_df_tranche, priority_count_tranche, on='PRIORITY', how='left')

# Fill missing values in Attribute Count with 0
full_combined_df_tranche['Attribute Count'] = full_combined_df_tranche['Attribute Count'].fillna(0)

# Set Month/Year to be the same for all MSG_TYP values
full_combined_df_tranche['Month/Year'] = previous_month_name

# Define banking and trading book types
banking_book_types_tranche = ["Business Finance", "Paragon"]
trading_book_types_tranche = ["3D Static", "Bloomberg", "Intex", "STS", "TAS"]

full_combined_df_tranche['Message Type Group'] = full_combined_df_tranche['MSG_TYP'].apply(
    lambda x: "Banking Book" if x in banking_book_types_tranche else "Trading Book" if x in trading_book_types_tranche else "Other"
)

# Add 'Exception trend' and 'Total' columns
full_combined_df_tranche['Exception trend'] = 'Deals'
full_combined_df_tranche['Total'] = total_value_tranche

# Reorder and select relevant columns as per the specified order
final_df_tranche = full_combined_df_tranche[['Exception trend', 'MSG_TYP', 'Message Type Group', 'Attribute Count', 
                             'PRIORITY', 'Volume', 'Total', 'Month/Year', 'STATUS', 'Priority Count']]

# Combine final_df and final_df_tranche into one dataframe
combined_final_df = pd.concat([final_df, final_df_tranche], axis=0)

# Define the path for the combined output Excel file
combined_output_file_name = f"ExceptionTrends_{previous_month_name}_script_output.xlsx"

combined_output_file_path = f"G:/girisuc/DQ CDE's Metrics and FDW Securitization/Exception_Trends_Monthly/{combined_output_file_name}"

combined_final_df.to_excel(combined_output_file_path, index=False)
print(f"Combined final_df and final_df_tranche saved to {combined_output_file_name}")

# Wait for 10 seconds
time.sleep(20)

# Define the folder path where the Excel files are located
folder_path = r"G:/girisuc/DQ CDE's Metrics and FDW Securitization/Exception_Trends_Monthly/"

# Use glob to get all the excel files in the folder
excel_files = glob.glob(folder_path + '*.xlsx')

if not excel_files:
    print("No Excel files found in the directory")
else:
    # Initialize an empty list to hold dataframes
    all_dataframes = []

    # Loop through the Excel files and append them to the list
    for file in excel_files:
        try:
            print(f"Processing file: {file} ")
            _, file_extension = os.path.splitext(file)
            if file_extension == '.xlsx':
                df = pd.read_excel(file, engine = 'openpyxl')
            elif file_extension == '.xls':
                df = pd.read_excel(file, engine = 'xlrd')
            else:
                continue  # Skip non-Excel files
            all_dataframes.append(df)
        except Exception as e:
            print(f"An error occured with file: {file}. Error: {e}")

if all_dataframes:
    # Concatenate all dataframes in the list
    combined_excel_df = pd.concat(all_dataframes, ignore_index=True)

    def convert_month_year(my):
        try:
            return pd.to_datetime(my, format='%b-%y', errors='coerce')
        except ValueError:
            try:
                return pd.to_datetime(my, format='%m/%d/%Y', errors='coerce')
            except ValueError:
                return pd.NaT
            
    # Apply the conversion function to the 'Month/Year' column
    combined_excel_df['Month/Year'] = combined_final_df['Month/Year'].apply(convert_month_year)

    # Sort the DataFrame by the 'Month/Year' column
    combined_excel_df.sort_values(by='Month/Year', inplace=True)

    # Define the path for the output Excel file
    output_file_path = r"G:/girisuc/DQ CDE's Metrics and FDW Securitization/Tableau REF/ExceptionTrends Tableau Input/ExceptionTrends_UAT.xlsx" 

    # Save the combined dataframe from all Excel files into one file
    combined_excel_df.to_excel(output_file_path, index=False)
    print(f"All Excel files combined and saved to {output_file_path}")
else:
    print("No dataframe were created from the file.")