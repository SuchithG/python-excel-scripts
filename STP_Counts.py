import pandas as pd
from datetime import datetime, timedelta

# Get the current date
current_date = datetime.now()

# Get the first day of the current month
first_day_of_current_month = current_date.replace(day=1)

# Calculate the last day of the previous month
last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)

# Get the month and year of the previous month
previous_month = last_day_of_previous_month.strftime('%b_%Y')

# Construct the file name
file_name = f"iRDS_SQL_Query_{previous_month}.xlsx"

# Load the Excel file
xls = pd.ExcelFile(file_name)

# Assuming the first sheet is 'Table1' containing the data for the first table
# Change these names according to your Excel file
table1 = pd.read_excel(xls, 'Table1')

# Calculate 'Total Nov exception' by summing up values across columns
total_nov_exception = table1[['Equity', 'Loans', 'LD', 'FI']].sum(axis=1)

# Add 'Total' column to the table
table1['Total'] = total_nov_exception

# Display the modified table with the 'Total' column
print(table1)
