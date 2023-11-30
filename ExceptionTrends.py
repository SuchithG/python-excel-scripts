import pandas as pd
from itertools import product
from datetime import datetime, timedelta

# Define previous month details
previous_month_date = (datetime.now().replace(day=1) - timedelta(days=1))
previous_month_name = previous_month_date.strftime("%b_%Y").upper()

# Load the data from the Excel file
file_path = 'path_to_your_excel_file.xlsx'  # Replace with the actual path
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

# Define the path for the output Excel file
output_file_path = 'output_data_analysis.xlsx'  # Replace with your desired file path

# Save the final dataframe to an Excel file
final_df.to_excel(output_file_path, index=False)

print(f"Data saved to {output_file_path}")
