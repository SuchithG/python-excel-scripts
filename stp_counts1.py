import pandas as pd
from datetime import datetime

# Define the age categories
age_categories = ['0-1 New', '02-07 days', '08-15 days', '16-30 days', '31-180 days', '>180 days']

# Define the function to determine the age category
def determine_age_category(creation_date, current_date):
    # If creation_date is a list, use the first element
    if isinstance(creation_date, list):
        creation_date = creation_date[0] if creation_date else None

    # Convert to datetime if it's a string or not a datetime object
    if isinstance(creation_date, str) or not isinstance(creation_date, datetime):
        try:
            creation_date = pd.to_datetime(creation_date)
        except Exception as e:
            raise ValueError(f"Error converting creation_date to datetime: {e}")

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
            sheet_data['TRUNC(NOTFCN_CRTE_TMS)'] = pd.to_datetime(sheet_data['TRUNC(NOTFCN_CRTE_TMS)'])
            sheet_data['TRUNC(LST_NOTFCN_TMS)'] = pd.to_datetime(sheet_data['TRUNC(LST_NOTFCN_TMS)'])

            open_records = sheet_data[sheet_data['NOTFCN_STAT_TYP'] == 'OPEN']
            closed_records = sheet_data[sheet_data['NOTFCN_STAT_TYP'] == 'CLOSED']

            open_records_combined = pd.concat([open_records_combined, open_records]).drop_duplicates()
            closed_records_combined = pd.concat([closed_records_combined, closed_records])

        closed_records_filtered = closed_records_combined[
            (closed_records_combined['TRUNC(LST_NOTFCN_TMS)'].dt.month == current_date.month) &
            (closed_records_combined['TRUNC(LST_NOTFCN_TMS)'].dt.year == current_date.year)
        ]

        final_combined_records = pd.concat([open_records_combined, closed_records_filtered])

        for _, row in final_combined_records.iterrows():
            creation_date = row['TRUNC(NOTFCN_CRTE_TMS)']
            if isinstance(creation_date, list):
                creation_date = creation_date[0] if creation_date else None
            
            if creation_date is None:
                continue

            age_category = determine_age_category(creation_date, current_date)
            count_column = 'COUNT(*)' if 'COUNT(*)' in row else 'COUNT(\'*\')'
            results_df.at[age_category, category] += row[count_column]

        print(f"Completed processing for category: {category}\n")

    return results_df

# Current date for processing
current_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
print(f"Current date for processing: {current_date}")

# Categories and their respective sheets
categories = {
    'Loans': ['Line 270', 'Line 297', 'Line 441', 'Line 523'],
    #'FI': ['Line 1616', 'Line 1407', 'Line 1727', 'Line 1843'],
    #'Equity': ['Line 764', 'Line 809', 'Line 970', 'Line 1024', 'Line 1088'],
    #'LD': ['Line 2104', 'Line 2261', 'Line 2325', 'Line 2389']
}

file_path = 'C:/Users/Suchith G/Documents/Test Docs/stp_counts.xlsx'  # Replace with your actual file path

# Process the file and create a DataFrame with the results
results_df = process_excel(file_path, categories, current_date)

# Display the final DataFrame
print("Final Results:")
print(results_df)
