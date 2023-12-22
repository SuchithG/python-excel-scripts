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
    results_df = pd.DataFrame(index=age_categories, columns=categories.keys()).fillna(0)

    # Process each category and its sheets
    for category, sheets in categories.items():
        print(f"Processing category: {category}")

        open_records_combined = pd.DataFrame()
        closed_records_combined = pd.DataFrame()

        for sheet_name in sheets:
            print(f"  Reading data from sheet: {sheet_name}")
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            sheet_data['TRUNC(NOTFCN_CRTE_TMS)'] = pd.to_datetime(sheet_data['TRUNC(NOTFCN_CRTE_TMS)'], errors='coerce')
            sheet_data['TRUNC(LST_NOTFCN_TMS)'] = pd.to_datetime(sheet_data['TRUNC(LST_NOTFCN_TMS)'], errors='coerce')

            # Separate OPEN and CLOSED records
            open_records = sheet_data[sheet_data['NOTFCN_STAT_TYP'] == 'OPEN']
            closed_records = sheet_data[sheet_data['NOTFCN_STAT_TYP'] == 'CLOSED']

            print(f"Found {len(open_records)} OPEN records in sheet: {sheet_name}")
            print(f"Found {len(closed_records)} CLOSED records in sheet: {sheet_name}")

            open_records_combined = pd.concat([open_records_combined, open_records]).drop_duplicates(subset='NOTFCN_ID')
            closed_records_combined = pd.concat([closed_records_combined, closed_records])

        print(f"Total OPEN records combined: {len(open_records_combined)}")
        print(f"Total CLOSED records combined before filtering: {len(closed_records_combined)}")

        # Filter CLOSED records based on last notification in the current month
        closed_records_filtered = closed_records_combined[
            (closed_records_combined['TRUNC(LST_NOTFCN_TMS)'].dt.month == current_date.month) &
            (closed_records_combined['TRUNC(LST_NOTFCN_TMS)'].dt.year == current_date.year)
        ]

        print(f"Total CLOSED records after filtering: {len(closed_records_filtered)}")

        final_combined_records = pd.concat([open_records_combined, closed_records_filtered])

        # Calculate counts for each unique row
        for _, row in final_combined_records.iterrows():
            age_category = determine_age_category(row['TRUNC(NOTFCN_CRTE_TMS)'], current_date)
            if age_category:
                count_column = 'COUNT(*)' if 'COUNT(*)' in row.index else 'COUNT(\'*\')'
                results_df.loc[age_category, category] += row[count_column]

        print(f"Completed processing for category: {category}\n")

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
