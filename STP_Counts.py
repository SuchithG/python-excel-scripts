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
def determine_age_category(creation_date, last_day_previous_month):
    age_days = (last_day_previous_month - creation_date).days
    for category, max_days in age_categories.items():
        if age_days <= max_days:
            return category
    return '>180 days'  # Default for any case not covered above

def process_excel_custom(file_path, categories, closed_sheets):
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
            try:
                # Read the sheet data
                sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)

                # Standardize the count column name
                if 'COUNT(*)' in sheet_data.columns or "COUNT('*')" in sheet_data.columns:
                    sheet_data.rename(columns={'COUNT(*)': 'Count', "COUNT('*')": 'Count'}, inplace=True)

                # Concatenate OPEN and CLOSED records
                open_records = sheet_data[sheet_data['NOTFCN_STAT_TYP'] == 'OPEN']
                closed_records = sheet_data[(sheet_data['NOTFCN_STAT_TYP'] == 'CLOSED') &
                                            (sheet_data['TRUNC(LST_NOTFCN_TMS)'] >= current_month_start) &
                                            (sheet_data['TRUNC(LST_NOTFCN_TMS)'] <= current_month_end)]
                category_open_records = pd.concat([category_open_records, open_records], ignore_index=True)
                category_closed_records = pd.concat([category_closed_records, closed_records], ignore_index=True)

            except Exception as e:
                print(f"Error processing sheet {sheet_name}: {e}")

        # Remove duplicates and calculate ageing for OPEN records
        category_open_records.drop_duplicates(subset=['TRUNC(NOTFCN_CRTE_TMS)', 'TRUNC(LST_NOTFCN_TMS)', 'NOTFCN_ID', 'NOTFCN_STAT_TYP'], inplace=True)
        for _, row in category_open_records.iterrows():
            creation_date = row['TRUNC(NOTFCN_CRTE_TMS)']
            if pd.notnull(creation_date):
                age_category = determine_age_category(creation_date, last_day_previous_month)
                count = pd.to_numeric(row['Count'], errors='coerce')
                open_ageing_df.at[age_category, category] += count

        # Remove duplicates and calculate ageing for CLOSED records
        category_closed_records.drop_duplicates(subset=['TRUNC(NOTFCN_CRTE_TMS)', 'TRUNC(LST_NOTFCN_TMS)', 'NOTFCN_ID', 'NOTFCN_STAT_TYP'], inplace=True)
        for _, row in category_closed_records.iterrows():
            creation_date = row['TRUNC(NOTFCN_CRTE_TMS)']
            if pd.notnull(creation_date):
                age_category = determine_age_category(creation_date, last_day_previous_month)
                count = pd.to_numeric(row['Count'], errors='coerce')
                closed_ageing_df.at[age_category, category] += count

    # Combine OPEN and CLOSED ageing DataFrames
    total_ageing_df = open_ageing_df.add(closed_ageing_df, fill_value=0)

    # Initialize the DataFrame for total exceptions
    total_exceptions_df = pd.DataFrame(index=['Open/Assign', 'Closed'], columns=categories.keys()).fillna(0)
    
    # Sum all values for Open/Assign in the total_ageing_df for each category
    for category in categories.keys():
        total_exceptions_df.loc['Open/Assign', category] = total_ageing_df[category].sum()
    
    # Read each sheet for Closed values and sum the 'COUNT(*)' column
    for category, sheet_name in closed_sheets.items():
        try:
            closed_data = pd.read_excel(file_path, sheet_name=sheet_name)
            total_exceptions_df.loc['Closed', category] = closed_data['COUNT(*)'].sum()
        except Exception as e:
            print(f"Error processing sheet {sheet_name} for closed counts: {e}")
    
    bulk_values = {'Equity': 0, 'Loans': 0, 'LD': 0, 'FI': 0}
    manual_values = {'Equity': 0, 'Loans': 0, 'LD': 0, 'FI': 0}
    auto_values = {'Equity': 0, 'Loans': 0, 'LD': 0, 'FI': 0}

    # Initialize the DataFrame for total breakup
    total_breakup_df = pd.DataFrame(index=['Bulk', 'Manual', 'Auto', 'Open'], columns=categories.keys())

    # Fill in the values for Bulk, Manual, Auto, and Open
    total_breakup_df.loc['Bulk'] = pd.Series(bulk_values)
    total_breakup_df.loc['Manual'] = pd.Series(manual_values)
    total_breakup_df.loc['Auto'] = pd.Series(auto_values)
    total_breakup_df.loc['Open'] = total_exceptions_df.loc['Open/Assign']

    return open_ageing_df, closed_ageing_df, total_ageing_df, total_exceptions_df, total_breakup_df

categories = {
    'Equity': ['Line 764', 'Line 809', 'Line 970', 'Line 1024', 'Line 1088']
}

closed_sheets = {
    'Equity': 'Line 655',
    'Loans': 'Line 180',
    'LD': 'Line 2020',
    'FI': 'Line 1280'
}

file_path = 'C:/Users/Suchith G/Documents/Test Docs/stp_counts.xlsx'

open_ageing_df, closed_ageing_df, total_ageing_df, total_exceptions_df, total_breakup_df  = process_excel_custom(file_path, categories, closed_sheets)

print("Open Ageing DataFrame:")
print(open_ageing_df)
print("\nClosed Ageing DataFrame:")
print(closed_ageing_df)
print("\nTotal Ageing DataFrame:")
print(total_ageing_df)
print("\nTotal exceptions DataFrame:")
print(total_exceptions_df)
print("\nTotal exceptions DataFrame:")
print(total_breakup_df)