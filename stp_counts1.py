import pandas as pd
from datetime import datetime

# Define the age categories
age_categories = ['0-1 New', '02-07 days', '08-15 days', '16-30 days', '31-180 days', '>180 days']

# Define the function to determine the age category
def determine_age_category(creation_date, current_date):
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
    # Initialize a dictionary to hold the sum of COUNT(*) for each category and age category
    sums_by_category_and_age = {cat: {age: 0 for age in age_categories} for cat in categories}

    # Process each category and its sheets
    for category, sheets in categories.items():
        print(f"Processing category: {category}")
        # Initialize a set to keep track of processed NOTFCN_IDs to avoid double-counting within the category
        processed_notfcn_ids = set()

        # Process each sheet within the category
        for sheet_name in sheets:
            print(f"  Reading data from sheet: {sheet_name}")
            # Load data from the sheet
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)

            # Normalize date columns to datetime
            sheet_data['TRUNC(NOTFCN_CRTE_TMS)'] = pd.to_datetime(sheet_data['TRUNC(NOTFCN_CRTE_TMS)'])
            sheet_data['TRUNC(LST_NOTFCN_TMS)'] = pd.to_datetime(sheet_data['TRUNC(LST_NOTFCN_TMS)'])

            # Determine the correct count column name
            count_column = 'COUNT(*)' if 'COUNT(*)' in sheet_data.columns else 'COUNT(\'*\')'

            # Process each row
            for index, row in sheet_data.iterrows():
                notfcn_id = row['NOTFCN_ID']
                notfcn_stat_typ = row['NOTFCN_STAT_TYP']
                count = row[count_column]

                # Check if the NOTFCN_ID has been processed already within the category
                if notfcn_id in processed_notfcn_ids:
                    continue

                # For OPEN records or CLOSED records with last notification in the current month, determine the age category
                if notfcn_stat_typ == 'OPEN' or (notfcn_stat_typ == 'CLOSED' and 
                                                 row['TRUNC(LST_NOTFCN_TMS)'].month == current_date.month and 
                                                 row['TRUNC(LST_NOTFCN_TMS)'].year == current_date.year):
                    age_category = determine_age_category(row['TRUNC(NOTFCN_CRTE_TMS)'], current_date)
                    sums_by_category_and_age[category][age_category] += count
                    processed_notfcn_ids.add(notfcn_id)

        print(f"Completed processing for category: {category}\n")

    return sums_by_category_and_age

# Assuming current_date is the first day of the current month
current_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
print(f"Current date for processing: {current_date}")

# Define the categories and their respective sheets
categories = {
    'Loans': ['Line 270', 'Line 297', 'Line 441', 'Line 447', 'Line 523'],
    #'FI': ['Line 1616', 'Line 1407', 'Line 1727', 'Line 1843'],
    #'Equity': ['Line 764', 'Line 809', 'Line 970', 'Line 1024', 'Line 1088'],
    #'LD': ['Line 2104', 'Line 2261', 'Line 2325', 'Line 2389']
}

file_path = 'C:/Users/Suchith G/Documents/Test Docs/stp_counts.xlsx'  # Replace with your actual file path

# Process the file and get the counts for each category
sums_by_category_and_age = process_excel(file_path, categories, current_date)

# Print the results in a user-friendly format
print("Summary of counts by category and age bracket:")
for category, age_counts in sums_by_category_and_age.items():
    print(f"Category: {category}")
    for age, count in age_counts.items():
        print(f"  {age}: {count}")
    print()  # Add an empty line for better readability between categories
