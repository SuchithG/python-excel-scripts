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
    results_df = pd.DataFrame(index=['0-1 New', '02-07 days', '08-15 days', '16-30 days', '31-180 days', '>180 days'], columns=categories.keys()).fillna(0)
    total_breakup_df = pd.DataFrame(index=['Bulk', 'Manual', 'Auto', 'Open'], columns=categories.keys()).fillna(0)
    
    for category, sheets in categories.items():
        all_records = pd.DataFrame()

        for sheet_name in sheets:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            count_col = 'COUNT(*)' if 'COUNT(*)' in sheet_data.columns else "COUNT('*')" if "COUNT('*')" in sheet_data.columns else None
            if count_col is None:
                continue  # Skip the sheet if no count column is found

            sheet_data['TRUNC(NOTFCN_CRTE_TMS)'] = pd.to_datetime(sheet_data['TRUNC(NOTFCN_CRTE_TMS)'], errors='coerce')
            all_records = pd.concat([all_records, sheet_data])

        all_records.drop_duplicates(inplace=True)

        # Calculate 'Open/Assign' and 'Closed' for summary DataFrame
        # This will be done after this loop

        # Calculate ageing counts for each category
        for _, row in all_records.iterrows():
            creation_date = row['TRUNC(NOTFCN_CRTE_TMS)']
            last_notification_date = row['TRUNC(LST_NOTFCN_TMS)']
            notification_status = row['NOTFCN_STAT_TYP']
            count = pd.to_numeric(row[count_col], errors='coerce')
            count = 0 if pd.isna(count) else count

            if pd.notnull(creation_date) and notification_status == 'OPEN':
                age_category = determine_age_category(creation_date, current_date)
                results_df.at[age_category, category] += count
            elif pd.notnull(creation_date) and notification_status == 'CLOSED' and pd.notnull(last_notification_date):
                if current_date.month == last_notification_date.month and current_date.year == last_notification_date.year:
                    age_category = determine_age_category(creation_date, current_date)
                    results_df.at[age_category, category] += count

        # Calculate 'Manual' entries for total breakup DataFrame
        if 'NOTFCN_ID' in all_records.columns and count_col:
            manual_filter = all_records['NOTFCN_ID'].astype(str).str.endswith('@db.com')
            manual_records = all_records[manual_filter]
            total_breakup_df.at['Manual', category] = manual_records[count_col].sum()

    # Now let's create the summary DataFrame like the first table in the image
    summary_df = pd.DataFrame(index=['Open/Assign', 'Closed'], columns=categories.keys())

    # Calculate 'Open/Assign' values by summing across ageing breaks
    for category in categories.keys():
        summary_df.at['Open/Assign', category] = results_df[category].sum()

    # Calculate 'Closed' value by summing 'COUNT(*)' from the specified sheets for each category
    closed_counts = {}
    for category, closed_sheet in summary_sheets.items():
        closed_data = pd.read_excel(file_path, sheet_name=closed_sheet)
        count_col = 'COUNT(*)' if 'COUNT(*)' in closed_data.columns else "COUNT('*')" if "COUNT('*')" in closed_data.columns else None
        if count_col:
            closed_counts[category] = closed_data[count_col].sum()
        else:
            closed_counts[category] = 0  # If no count column, set to 0
    
    for category in categories.keys():
        summary_df.at['Closed', category] = closed_counts.get(category, 0)

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
