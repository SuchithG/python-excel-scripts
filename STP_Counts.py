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

def process_excel_custom(file_path, categories, current_date):
    results_df = pd.DataFrame(index=['0-1 New', '02-07 days', '08-15 days', '16-30 days', '31-180 days', '>180 days'], columns=categories.keys()).fillna(0)

    for category, sheets in categories.items():
        all_records = pd.DataFrame()

        for sheet_name in sheets:
            cols_to_read = ["TRUNC(NOTFCN_CRTE_TMS)", "TRUNC(LST_NOTFCN_TMS)", "NOTFCN_ID", "NOTFCN_STAT_TYP", "COUNT(*)"]
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name, usecols=cols_to_read)

            # Convert columns to datetime
            sheet_data['TRUNC(NOTFCN_CRTE_TMS)'] = pd.to_datetime(sheet_data['TRUNC(NOTFCN_CRTE_TMS)'], errors='coerce')
            sheet_data['TRUNC(LST_NOTFCN_TMS)'] = pd.to_datetime(sheet_data['TRUNC(LST_NOTFCN_TMS)'], errors='coerce')

            all_records = pd.concat([all_records, sheet_data])

        all_records.drop_duplicates(inplace=True)

        for _, row in all_records.iterrows():
            creation_date = row['TRUNC(NOTFCN_CRTE_TMS)']
            last_notification_date = row['TRUNC(LST_NOTFCN_TMS)']
            notification_status = row['NOTFCN_STAT_TYP']
            count = pd.to_numeric(row['COUNT(*)'], errors='coerce')
            count = 0 if pd.isna(count) else count

            if pd.notnull(creation_date) and notification_status == 'OPEN':
                age_category = determine_age_category(creation_date, current_date)
                results_df.at[age_category, category] += count
            elif pd.notnull(creation_date) and notification_status == 'CLOSED' and pd.notnull(last_notification_date):
                if current_date.month == last_notification_date.month and current_date.year == last_notification_date.year:
                    age_category = determine_age_category(creation_date, current_date)
                    results_df.at[age_category, category] += count

    return results_df

# Example usage
current_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
categories = {
    'Equity': ['Line 764', 'Line 809', 'Line 970', 'Line 1024', 'Line 1088']
}
file_path = 'C:/Users/Suchith G/Documents/Test Docs/stp_counts.xlsx'  # Update this with your file path

# Process the file and create a DataFrame with the results
results_df = process_excel_custom(file_path, categories, current_date)

# Display the DataFrame
print(results_df)
