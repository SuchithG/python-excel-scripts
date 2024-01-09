import pandas as pd
from datetime import datetime

# Define the age categories as per the new logic
def determine_age_category(creation_date, current_date):
    previous_month_end = current_date.replace(day=1) - pd.Timedelta(days=1)
    previous_month_start = previous_month_end.replace(day=1)

    if previous_month_start <= creation_date <= previous_month_end:
        day_of_month = creation_date.day
        if day_of_month >= 30:  # 30th and 31st
            return '0-1 New'
        elif day_of_month >= 25:  # 25th to 29th
            return '02-07 days'
        elif day_of_month >= 16:  # 16th to 24th
            return '08-15 days'
        else:  # First 15 days
            return '16-30 days'
    elif creation_date < previous_month_start - pd.Timedelta(days=180):
        return '>180 days'
    else:  # Previous month and up to 180 days
        return '31-180 days'

def process_excel_custom(file_path, categories, current_date, summary_sheets):
    # Define the start and end of the current month
    current_month_start = current_date.replace(day=1)
    current_month_end = current_date.replace(day=1, month=current_date.month % 12 + 1) - pd.Timedelta(days=1)

    results_df = pd.DataFrame(index=['0-1 New', '02-07 days', '08-15 days', '16-30 days', '31-180 days', '>180 days'], columns=categories.keys()).fillna(0)
    total_breakup_df = pd.DataFrame(index=['Bulk', 'Manual', 'Auto', 'Open'], columns=categories.keys()).fillna(0)

    for category, sheets in categories.items():
        category_open_records = pd.DataFrame()
        category_closed_records = pd.DataFrame()

        for sheet_name in sheets:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            count_col = 'COUNT(*)' if 'COUNT(*)' in sheet_data.columns else "COUNT('*')" if "COUNT('*')" in sheet_data.columns else None
            if count_col is None:
                continue  # Skip the sheet if no count column is found

            sheet_data['TRUNC(NOTFCN_CRTE_TMS)'] = pd.to_datetime(sheet_data['TRUNC(NOTFCN_CRTE_TMS)'], errors='coerce')
            
            # Split OPEN and CLOSED records and filter
            open_records = sheet_data[sheet_data['NOTFCN_STAT_TYP'] == 'OPEN']
            closed_records = sheet_data[(sheet_data['NOTFCN_STAT_TYP'] == 'CLOSED') & 
                                        (sheet_data['TRUNC(LST_NOTFCN_TMS)'] >= current_month_start) & 
                                        (sheet_data['TRUNC(LST_NOTFCN_TMS)'] <= current_month_end)]

            # Append to the category-specific DataFrames
            category_open_records = pd.concat([category_open_records, open_records])
            category_closed_records = pd.concat([category_closed_records, closed_records])

        # Combine OPEN and CLOSED records for the category
        category_combined_records = pd.concat([category_open_records, category_closed_records])

        # Remove duplicates across all columns
        category_combined_records.drop_duplicates(inplace=True)

        # Calculate ageing counts for the category
        for _, row in category_combined_records.iterrows():
            creation_date = row['TRUNC(NOTFCN_CRTE_TMS)']
            count = pd.to_numeric(row[count_col], errors='coerce')
            count = 0 if pd.isna(count) else count

            if pd.notnull(creation_date):
                age_category = determine_age_category(creation_date, current_date)
                results_df.at[age_category, category] += count

    # Calculate 'Manual', 'Bulk', 'Auto', and 'Open' for total_breakup_df
    # ... (additional logic for 'Manual', 'Bulk', 'Auto', and 'Open')

    # Calculate summary_df
    summary_df = pd.DataFrame(index=['Open/Assign', 'Closed'], columns=categories.keys()).fillna(0)
    
    # Populate summary_df with counts
    for category in categories.keys():
        summary_df.at['Open/Assign', category] = category_open_records[count_col].sum()
        summary_df.at['Closed', category] = category_closed_records[count_col].sum()

    return results_df, summary_df, total_breakup_df
        
# Example usage
current_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)

categories = {
    'Equity': ['Line 764', 'Line 809', 'Line 970', 'Line 1024', 'Line 1088'],
    'Loans': ['Line 270', 'Line 297', 'Line 441', 'Line 447', 'Line 523'],
    'FI': ['Line 1616', 'Line 1407', 'Line 1727', 'Line 1843'],
    'LD': ['Line 2104', 'Line 2261', 'Line 2325', 'Line 2389']
}

summary_sheets = {
    'Equity': 'Line 655',
    'Loans': 'Line 180',
    'FI': 'Line 1280',
    'LD': 'Line 2020'
}

file_path = 'C:/Users/Suchith G/Documents/Test Docs/stp_counts.xlsx'  # Update this with your file path

# Process the file and create DataFrames with the results
results_df, summary_df, total_breakup_df  = process_excel_custom(file_path, categories, current_date, summary_sheets)

# Display the DataFrames
print("Ageing DataFrame:")
print(results_df)
print("\nSummary DataFrame:")
print(summary_df)
print("\nTotal Breakup DataFrame:")
print(total_breakup_df)
