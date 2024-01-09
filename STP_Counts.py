import pandas as pd
from datetime import datetime

# Define the age categories
age_categories = {
    '0-1 New': 1,
    '02-07 days': 7,
    '08-15 days': 15,
    '16-30 days': 30,
    '31-180 days': 180,
    '>180 days': float('inf')
}

# Define the function to determine the age category
def determine_age_category(creation_date, current_date):
    age_days = (current_date - creation_date).days
    for category, max_days in age_categories.items():
        if age_days <= max_days:
            return category
    return '>180 days'  # Default for any case not covered above

def process_excel_custom(file_path, categories):
    # Define the current date for ageing calculation as the last day of the previous month
    current_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_day_previous_month = current_date - pd.Timedelta(days=1)
    current_month_start = current_date.replace(day=1)
    current_month_end = current_date.replace(day=1, month=current_date.month % 12 + 1) - pd.Timedelta(days=1)

    # Initialize the DataFrames for open and closed ageing
    open_ageing_df = pd.DataFrame(index=age_categories.keys(), columns=categories.keys()).fillna(0)
    closed_ageing_df = pd.DataFrame(index=age_categories.keys(), columns=categories.keys()).fillna(0)

    # Process OPEN and CLOSED records for each category
    for category, sheets in categories.items():
        category_open_records = pd.DataFrame()
        category_closed_records = pd.DataFrame()
        
        for sheet_name in sheets:
            # Read the sheet data
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)

            # Concatenate OPEN records
            open_records = sheet_data[sheet_data['NOTFCN_STAT_TYP'] == 'OPEN']
            category_open_records = pd.concat([category_open_records, open_records], ignore_index=True)

            # Concatenate CLOSED records within the current month
            closed_records = sheet_data[(sheet_data['NOTFCN_STAT_TYP'] == 'CLOSED') &
                                        (sheet_data['TRUNC(LST_NOTFCN_TMS)'] >= current_month_start) &
                                        (sheet_data['TRUNC(LST_NOTFCN_TMS)'] <= current_month_end)]
            category_closed_records = pd.concat([category_closed_records, closed_records], ignore_index=True)

        # Remove duplicates and calculate ageing for OPEN records
        category_open_records.drop_duplicates(subset=['TRUNC(NOTFCN_CRTE_TMS)', 'TRUNC(LST_NOTFCN_TMS)', 'NOTFCN_ID', 'NOTFCN_STAT_TYP'], inplace=True)
        for _, row in category_open_records.iterrows():
            creation_date = row['TRUNC(NOTFCN_CRTE_TMS)']
            if pd.notnull(creation_date):
                age_category = determine_age_category(creation_date, last_day_previous_month)
                count = pd.to_numeric(row['COUNT(*)'], errors='coerce')
                open_ageing_df.at[age_category, category] += count

        # Remove duplicates and calculate ageing for CLOSED records
        category_closed_records.drop_duplicates(subset=['TRUNC(NOTFCN_CRTE_TMS)', 'TRUNC(LST_NOTFCN_TMS)', 'NOTFCN_ID', 'NOTFCN_STAT_TYP'], inplace=True)
        for _, row in category_closed_records.iterrows():
            creation_date = row['TRUNC(NOTFCN_CRTE_TMS)']
            if pd.notnull(creation_date):
                age_category = determine_age_category(creation_date, last_day_previous_month)
                count = pd.to_numeric(row['COUNT(*)'], errors='coerce')
                closed_ageing_df.at[age_category, category] += count

    return open_ageing_df, closed_ageing_df

# Example usage
categories = {
    'Equity': ['Line 764', 'Line 809', 'Line 970', 'Line 1024', 'Line 1088']
}
file_path = 'C:/Users/Suchith G/Documents/Test Docs/stp_counts.xlsx'  # Update this with your file path

# Process the file and create DataFrames with the results
open_ageing_df, closed_ageing_df = process_excel_custom(file_path, categories)

# Combine OPEN and CLOSED ageing DataFrames to create the total_ageing_df
total_ageing_df = open_ageing_df.add(closed_ageing_df, fill_value=0)

# Display the DataFrames
print("Open Ageing DataFrame:")
print(open_ageing_df)
print("\nClosed Ageing DataFrame:")
print(closed_ageing_df)
print("\nTotal Ageing DataFrame:")
print(total_ageing_df)
