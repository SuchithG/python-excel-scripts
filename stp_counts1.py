import pandas as pd
from datetime import datetime, timedelta

# Define the function to get the previous month's date range
def get_previous_month_range(current_date):
    first_day_of_current_month = current_date.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    first_day_of_previous_month = last_day_of_previous_month.replace(day=1)
    return first_day_of_previous_month, last_day_of_previous_month

# Define the function to categorize the notification age
def categorize_notification(created_date, first_day_of_previous_month, last_day_of_previous_month):
    if last_day_of_previous_month.day - 1 <= created_date.day <= last_day_of_previous_month.day:
        return '0-1 New'
    elif (24 <= created_date.day <= 28) or (last_day_of_previous_month.day == 31 and 25 <= created_date.day <= 29):
        return '02-07 days'
    elif (16 <= created_date.day <= 23) or (last_day_of_previous_month.day == 31 and 16 <= created_date.day <= 24):
        return '08-15 days'
    elif 1 <= created_date.day <= 15:
        return '16-30 days'
    elif (current_date - created_date).days <= 180:
        return '31-180 days'
    else:
        return '>180 days'

# Define the function to process each sheet's data
def process_sheet_data(sheet_data, current_date):
    sheet_data['TRUNC(NOTFCN_CRTE_TMS)'] = pd.to_datetime(sheet_data['TRUNC(NOTFCN_CRTE_TMS)'])
    first_day_of_previous_month, last_day_of_previous_month = get_previous_month_range(current_date)

    # Categorize notifications
    sheet_data['Age_Category'] = sheet_data['TRUNC(NOTFCN_CRTE_TMS)'].apply(
        lambda x: categorize_notification(x, first_day_of_previous_month, last_day_of_previous_month)
    )

    # Count notifications in each category
    category_counts = sheet_data['Age_Category'].value_counts().reset_index()
    category_counts.columns = ['Age_Category', 'Count']
    return category_counts

# Define the function to process all notifications and format the output
def process_notifications(file_path, sheets, current_date):
    all_counts = pd.DataFrame()

    for category, sheet_names in sheets.items():
        for sheet_name in sheet_names:
            sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
            sheet_counts = process_sheet_data(sheet_data, current_date)
            sheet_counts['Category'] = category
            all_counts = pd.concat([all_counts, sheet_counts])

    # Aggregate the counts for each category and age category
    aggregate_counts = all_counts.groupby(['Category', 'Age_Category']).sum().unstack(fill_value=0)
    aggregate_counts.columns = aggregate_counts.columns.droplevel(0)  # Drop the top level ('Count')

    # Ensure all expected age categories are present
    age_categories = ['0-1 New', '02-07 days', '08-15 days', '16-30 days', '31-180 days', '>180 days']
    for category in age_categories:
        if category not in aggregate_counts:
            aggregate_counts[category] = 0

    # Reorder the columns to match the expected age categories
    aggregate_counts = aggregate_counts[age_categories]

    # Reset index to turn the categories into a column and prepare for final formatting
    final_output = aggregate_counts.reset_index()

    return final_output

# Configuration: replace with the path to your Excel file and current date
file_path = 'path_to_your_file.xlsx'
current_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
sheets = {
    'Loans': ['Line 270', 'Line 297', 'Line 441', 'Line 523'],
    # Add other categories and sheet names as necessary
}

# Process the notifications and get the formatted output
formatted_output = process_notifications(file_path, sheets, current_date)

# Print and optionally save the formatted output to an Excel file
print(formatted_output)
# formatted_output.to_excel('formatted_notification_counts.xlsx', index=False)
