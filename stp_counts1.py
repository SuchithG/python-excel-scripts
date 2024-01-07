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

def process_excel(file_path, categories, current_date):
    # Determine the start and end of the current month
    current_month_start = current_date.replace(day=1)
    current_month_end = current_date.replace(month=current_date.month % 12 + 1, day=1) - pd.Timedelta(days=1)

    results_df = pd.DataFrame(index=age_categories.keys(), columns=categories.keys()).fillna(0)

    for category, sheets in categories.items():
        all_records = pd.DataFrame()

        for sheet_name in sheets:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            # Convert columns to datetime and handle errors
            sheet_data['TRUNC(NOTFCN_CRTE_TMS)'] = pd.to_datetime(sheet_data['TRUNC(NOTFCN_CRTE_TMS)'], errors='coerce')
            sheet_data['TRUNC(LST_NOTFCN_TMS)'] = pd.to_datetime(sheet_data['TRUNC(LST_NOTFCN_TMS)'], errors='coerce')

            # Append sheet data to the combined DataFrame
            all_records = pd.concat([all_records, sheet_data])

        # Drop duplicates across all columns for the combined data
        all_records.drop_duplicates(inplace=True)

        # Combine OPEN and CLOSED records for final processing
        final_records = all_records[all_records['NOTFCN_STAT_TYP'].isin(['OPEN', 'CLOSED'])]

        # Calculate counts for each unique row
        for _, row in final_records.iterrows():
            creation_date = row['TRUNC(NOTFCN_CRTE_TMS)']
            last_notification_date = row['TRUNC(LST_NOTFCN_TMS)']

            if pd.notnull(creation_date):
                # Apply additional filtering for 'CLOSED' status
                if row['NOTFCN_STAT_TYP'] == 'CLOSED' and pd.notnull(last_notification_date):
                    # Skip the row if the last notification date is not in the current month
                    if not (current_month_start <= last_notification_date <= current_month_end):
                        continue

                age_category = determine_age_category(creation_date, current_date)
                # Ensure the count column is numeric and handle NaN values
                count = pd.to_numeric(row.get('COUNT(*)', 0), errors='coerce')
                count = 0 if pd.isna(count) else count
                results_df.at[age_category, category] += count

    return results_df

# Example usage
current_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
print(f"Current date for processing: {current_date.strftime('%Y-%m-%d')}")

categories = {
    #'Loans': ['Line 270', 'Line 297', 'Line 441', 'Line 523']
    'Equity': ['Line 764', 'Line 809', 'Line 970', 'Line 1024', 'Line 1088']
}

# Replace the file path with your actual file path
file_path = 'C:/Users/Suchith G/Documents/Test Docs/stp_counts.xlsx'  # Update with the path to the uploaded file

# Process the file and create a DataFrame with the results
results_df = process_excel(file_path, categories, current_date)

# Display the final DataFrame
print("Final Results:")
print(results_df)
