import pandas as pd
from datetime import datetime

# Define the age categories
age_categories = ['0-1 New', '02-07 days', '08-15 days', '16-30 days', '31-180 days', '>180 days']

# Define the function to determine the age category
def determine_age_category(creation_date, current_date):
    if isinstance(creation_date, list) or isinstance(creation_date, pd.Series):
        creation_date = creation_date.iloc[0] if not creation_date.empty else None

    if isinstance(creation_date, str):
        try:
            creation_date = pd.to_datetime(creation_date)
        except Exception as e:
            print(f"Error converting creation_date to datetime: {e}, value: {creation_date}")
            return None

    age_days = (current_date - creation_date).days
    if age_days <= 1:
        return '0-1 New'
    elif 2 <= age_days <= 7:
        return '02-07 days'
    elif 8 <= age_days <= 15:
        return '08-15 days'
    elif 16 <= age_days <= 30:
        return '16-30 days'
    elif 31 <= age_days <= 180:
        return '31-180 days'
    else:
        return '>180 days'

# Define the function to process the Excel file
def process_excel(file_path, categories, current_date):
    results_df = pd.DataFrame(index=age_categories.keys(), columns=categories.keys()).fillna(0)

    for category, sheets in categories.items():
        all_records = pd.DataFrame()

        for sheet_name in sheets:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            all_records = pd.concat([all_records, sheet_data])

        # Drop duplicates across all columns for both OPEN and CLOSED records
        all_records.drop_duplicates(inplace=True)

        # Debugging: Print the records after deduplication
        print(f"Deduplicated records for {category}: {len(all_records)}")
        print(all_records)

        # Filter records into open and closed based on status and last notification time for closed
        open_records = all_records[all_records['NOTFCN_STAT_TYP'] == 'OPEN']
        closed_records = all_records[
            (all_records['NOTFCN_STAT_TYP'] == 'CLOSED') &
            (pd.to_datetime(all_records['TRUNC(LST_NOTFCN_TMS)'], errors='coerce').dt.month == current_date.month) &
            (pd.to_datetime(all_records['TRUNC(LST_NOTFCN_TMS)'], errors='coerce').dt.year == current_date.year)
        ]

        # Combine OPEN and CLOSED records for final processing
        final_records = pd.concat([open_records, closed_records])

        # Calculate counts for each unique row
        for _, row in final_records.iterrows():
            creation_date = pd.to_datetime(row['TRUNC(NOTFCN_CRTE_TMS)'], errors='coerce')
            age_category = determine_age_category(creation_date, current_date)
            if age_category:
                count_column_name = 'COUNT(*)' if 'COUNT(*)' in row else "COUNT('*')"
                count = row.get(count_column_name, 0)
                results_df.at[age_category, category] += count

                # Debugging: Print each record's details
                print(f"ID: {row.get('NOTFCN_ID', 'N/A')}, Age Category: {age_category}, Count: {count}")

    return results_df

# Current date for processing
current_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
print(f"Current date for processing: {current_date}")

# Categories and their respective sheets
categories = {
    'Loans': ['Line 270', 'Line 297', 'Line 441', 'Line 523']
}

# Replace the file path with your actual file path
file_path = 'C:/Users/Suchith G/Documents/Test Docs/stp_counts.xlsx'  # Update with the path to the uploaded file

# Process the file and create a DataFrame with the results
results_df = process_excel(file_path, categories, current_date)

# Display the final DataFrame
print("Final Results:")
print(results_df)
