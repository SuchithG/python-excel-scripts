import pandas as pd

# Load the Excel file
excel_file = 'your_file.xlsx'  # Replace 'your_file.xlsx' with your file name
xls = pd.ExcelFile(excel_file)

# Assuming the first sheet is 'Table1' containing the data for the first table
# Change these names according to your Excel file
table1 = pd.read_excel(xls, 'Table1')

# Calculate 'Total Nov exception' by summing up values across columns
total_nov_exception = table1[['Equity', 'Loans', 'LD', 'FI']].sum(axis=1)

# Add 'Total' column to the table
table1['Total'] = total_nov_exception

# Display the modified table with the 'Total' column
print(table1)
