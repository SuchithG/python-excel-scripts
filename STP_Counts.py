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
    # Determine the start and end of the current month
    current_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_day_previous_month = current_date - pd.Timedelta(days=1)
    current_month_start = current_date.replace(day=1)
    current_month_end = current_date.replace(month=current_date.month % 12 + 1, day=1) - pd.Timedelta(days=1)

    # Initialize the DataFrames for open and closed ageing
    open_ageing_df = pd.DataFrame(index=age_categories.keys(), columns=categories.keys()).fillna(0)
    closed_ageing_df = pd.DataFrame(index=age_categories.keys(), columns=categories.keys()).fillna(0)
    total_breakup_df = pd.DataFrame(index=['Bulk', 'Manual', 'Auto', 'Open'], columns=categories.keys()).fillna(0)
    total_exceptions_df = pd.DataFrame(index=['Open/Assign', 'Closed'], columns=categories.keys()).fillna(0)

    # Process OPEN and CLOSED records for each category
    for category, sheets in categories.items():
        category_open_records = pd.DataFrame()
        category_closed_records = pd.DataFrame()
        
        for sheet_name in sheets:
            try:
                sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
                sheet_data.rename(columns=lambda x: 'Count' if 'COUNT' in x else x, inplace=True)

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
                count = pd.to_numeric(row.get('Count', 0), errors='coerce')
                count = 0 if pd.isna(count) else count
                open_ageing_df.at[age_category, category] += count

        # Remove duplicates and calculate ageing for CLOSED records
        category_closed_records.drop_duplicates(subset=['TRUNC(NOTFCN_CRTE_TMS)', 'TRUNC(LST_NOTFCN_TMS)', 'NOTFCN_ID', 'NOTFCN_STAT_TYP'], inplace=True)
        for _, row in category_closed_records.iterrows():
            creation_date = row['TRUNC(NOTFCN_CRTE_TMS)']
            if pd.notnull(creation_date):
                age_category = determine_age_category(creation_date, last_day_previous_month)
                count = pd.to_numeric(row.get('Count', 0), errors='coerce')
                count = 0 if pd.isna(count) else count
                closed_ageing_df.at[age_category, category] += count

        # Ensure there are no NaN values before summing
        open_ageing_df.fillna(0, inplace=True)
        closed_ageing_df.fillna(0, inplace=True)

    # Sum up the total ageing data
    total_ageing_df = open_ageing_df.add(closed_ageing_df, fill_value=0)

    # Calculate 'Manual' counts for total_breakup_df using closed_sheets
    for category, sheet_name in closed_sheets.items():
        try:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            sheet_data.rename(columns=lambda x: 'Count' if 'COUNT' in x else x, inplace=True)
            manual_filter = sheet_data['LAST_CHG_USR_ID'].astype(str).str.endswith('@db.com')
            manual_counts = sheet_data.loc[manual_filter, 'Count'].sum()
            total_breakup_df.at['Manual', category] = manual_counts
        except Exception as e:
            print(f"Error processing sheet {sheet_name} for manual counts: {e}")

    # Calculate 'Open/Assign' counts for total_exceptions_df
    for category in categories.keys():
        total_exceptions_df.at['Open/Assign', category] = total_ageing_df[category].sum()

    # Calculate 'Closed' counts for total_exceptions_df using closed_sheets
    for category, sheet_name in closed_sheets.items():
        try:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            # Standardize the count column name
            if 'COUNT(*)' in sheet_data.columns:
                count_col = 'COUNT(*)'
            elif "COUNT('*')" in sheet_data.columns:
                count_col = "COUNT('*')"
            else:
                print(f"Count column not found in sheet {sheet_name}")
                continue

            # Sum the counts for closed records
            total_exceptions_df.at['Closed', category] = sheet_data[count_col].sum()
        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {e}")

        # Populate the 'Open' values for total_breakup_df
        for category in categories.keys():
            total_breakup_df.at['Open', category] = total_exceptions_df.at['Open/Assign', category]

    # Calculate 'Bulk' counts for total_breakup_df using closed_sheets
    for category, sheet_name in closed_sheets.items():
        try:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            # Standardize the count column name
            if 'COUNT(*)' in sheet_data.columns:
                count_col = 'COUNT(*)'
            elif "COUNT('*')" in sheet_data.columns:
                count_col = "COUNT('*')"
            else:
                print(f"Count column not found in sheet {sheet_name}")
                continue

            # Filter records for bulk counts
            bulk_filter = (sheet_data['LAST_CHG_USR_ID'].astype(str).str.endswith('.txt') |
                        sheet_data['LAST_CHG_USR_ID'].astype(str).str.contains('EX_CLS') |
                        sheet_data['LAST_CHG_USR_ID'].astype(str).str.startswith('INC'))
            bulk_counts = sheet_data.loc[bulk_filter, count_col].sum()
            total_breakup_df.at['Bulk', category] = bulk_counts
        except Exception as e:
            print(f"Error processing sheet {sheet_name} for bulk counts: {e}")

    # Define the MSG_TYP for each category
    msg_type = {
        'Loans': ['BBSyndicatedLoans', 'BBSyndicatedLoanBulk', 'BBSyndicatedLoans_S&P', 'BBSyndicatedLoans_Moodys'],
        'FI': ['cRDS_iRDS', 'BB_CorpPrefConvGovt_S&P'],
        'Equity': ['ASX_Security_Masters', 'BBEquityCollateral'],
        'LD': ['FOWSessionTime', 'FOWContracts']
    }

    # Calculate 'Auto' counts for total_breakup_df using closed_sheets and msg_type
    for category, sheet_name in closed_sheets.items():
        try:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            # Standardize the count column name
            sheet_data.rename(columns=lambda x: 'Count' if 'COUNT' in x else x, inplace=True)

            # Filter records for auto counts
            auto_filter = (sheet_data['LAST_CHG_USR_ID'].astype(str).str.contains('Auto', case=False) |
                        sheet_data['MSG_TYP'].isin(msg_type[category]))
            auto_counts = sheet_data.loc[auto_filter, 'Count'].sum()
            total_breakup_df.at['Auto', category] = auto_counts
        except Exception as e:
            print(f"Error processing sheet {sheet_name} for auto counts: {e}")

    # Calculate 'Auto' counts for total_breakup_df using closed_sheets and msg_type
    for category, sheet_name in closed_sheets.items():
        try:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Standardize the count column name to 'Count'
            if 'COUNT(*)' in sheet_data.columns:
                sheet_data.rename(columns={'COUNT(*)': 'Count'}, inplace=True)
            elif "COUNT('*')" in sheet_data.columns:
                sheet_data.rename(columns={"COUNT('*')": 'Count'}, inplace=True)

            # Ensure there are no NaN values before summing
            sheet_data['Count'].fillna(0, inplace=True)

            # Define the message types for each category
            msg_types = {
                'Loans': ['BBSyndicatedLoans', 'BBSyndicatedLoanBulk', 'BBSyndicatedLoans_S&P','BBSyndicatedLoans_Moodys'],
                'FI': ['cRDS_iRDS', 'BB_CorpPrefConvGovt_S&P'],
                'Equity': ['ASX_Security_Masters', 'BBEquityCollateral'],
                'LD': ['FOWSessionTime','FOWContracts'] 
            }

            # Filter records for auto counts
            auto_filter = (sheet_data['LAST_CHG_USR_ID'].astype(str).str.contains('Auto', case=False) |
                        sheet_data['LAST_CHG_USR_ID'].astype(str).str.endswith('.txt') |
                        sheet_data['MSG_TYP'].isin(msg_types.get(category, [])))
            auto_counts = sheet_data.loc[auto_filter, 'Count'].sum()
            total_breakup_df.at['Auto', category] = auto_counts
        except Exception as e:
            print(f"Error processing sheet {sheet_name} for auto counts: {e}")

    # Return the completed DataFrames
    return open_ageing_df, closed_ageing_df, total_ageing_df, total_exceptions_df, total_breakup_df

# Example usage
categories = {
    'Equity': ['Line 764', 'Line 809', 'Line 970', 'Line 1024', 'Line 1088'],
    'Loans': ['Line 270', 'Line 297', 'Line 441', 'Line 447', 'Line 523'],
    'LD': ['Line 2104', 'Line 2261', 'Line 2325', 'Line 2389'],
    'FI': ['Line 1616', 'Line 1407', 'Line 1727', 'Line 1843']
}

closed_sheets = {
    'Equity': 'Line 655',
    'Loans': 'Line 180',
    'LD': 'Line 2020',
    'FI': 'Line 1280'
}

file_path = 'C:/Users/Suchith G/Documents/Test Docs/stp_counts.xlsx'

# Execute the processing function
open_ageing_df, closed_ageing_df, total_ageing_df, total_exceptions_df, total_breakup_df = process_excel_custom(file_path, categories, closed_sheets)

# Display the DataFrames
print("Open Ageing DataFrame:")
print(open_ageing_df)
print("\nClosed Ageing DataFrame:")
print(closed_ageing_df)
print("\nTotal Ageing DataFrame:")
print(total_ageing_df)
print("\nTotal Exceptions DataFrame:")
print(total_exceptions_df)
print("\nTotal Breakup DataFrame:")
print(total_breakup_df)
