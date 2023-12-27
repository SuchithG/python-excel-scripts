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

# Define the function to process the Excel file
def process_excel(file_path, categories, current_date):
    results_df = pd.DataFrame(index=age_categories.keys(), columns=categories.keys()).fillna(0)

    for category, sheets in categories.items():
        all_records = pd.DataFrame()

        for sheet_name in sheets:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            # Convert columns to datetime
            sheet_data['TRUNC(NOTFCN_CRTE_TMS)'] = pd.to_datetime(sheet_data['TRUNC(NOTFCN_CRTE_TMS)'], errors='coerce')
            sheet_data['TRUNC(LST_NOTFCN_TMS)'] = pd.to_datetime(sheet_data['TRUNC(LST_NOTFCN_TMS)'], errors='coerce')

            # Append sheet data to the combined DataFrame
            all_records = pd.concat([all_records, sheet_data])

        # Drop duplicates across all columns for the combined data
        all_records.drop_duplicates(inplace=True)

        # Separate OPEN and CLOSED records
        open_records = all_records[all_records['NOTFCN_STAT_TYP'] == 'OPEN']
        closed_records = all_records[all_records['NOTFCN_STAT_TYP'] == 'CLOSED']

        # Filter CLOSED records to include those with last notification time in the current month
        closed_records_filtered = closed_records[
            (closed_records['TRUNC(LST_NOTFCN_TMS)'].dt.month == current_date.month) &
            (closed_records['TRUNC(LST_NOTFCN_TMS)'].dt.year == current_date.year)
        ]

        # Debugging: Print the filtered CLOSED records and their count
        print(f"Filtered CLOSED records for {category}:")
        print(closed_records_filtered[['NOTFCN_ID', 'TRUNC(NOTFCN_CRTE_TMS)', 'TRUNC(LST_NOTFCN_TMS)', 'NOTFCN_STAT_TYP', 'COUNT(*)']])

        # Combine OPEN and filtered CLOSED records for final processing
        final_records = pd.concat([open_records, closed_records_filtered])

        # Calculate counts for each unique row
        for _, row in final_records.iterrows():
            creation_date = row['TRUNC(NOTFCN_CRTE_TMS)']
            if pd.notnull(creation_date):
                age_category = determine_age_category(creation_date, current_date)
                count_column_name = 'COUNT(*)' if 'COUNT(*)' in row else "COUNT('*')"
                count = row.get(count_column_name, 0)
                results_df.at[age_category, category] += count

                # Debugging: Print each record's details
                print(f"ID: {row.get('NOTFCN_ID', 'N/A')}, Age Category: {age_category}, Count: {count}")

    return results_df

# Current date for processing
current_date = datetime(2023, 12, 1)
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
