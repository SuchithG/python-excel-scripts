import pandas as pd

# Load the Excel file
excel_file = 'your_file.xlsx'  # Replace 'your_file.xlsx' with your file name
xls = pd.ExcelFile(excel_file)

# Assuming you have three sheets named 'Table1', 'Table2', and 'Table3'
# Change these names according to your Excel file
table1 = pd.read_excel(xls, 'Table1')
table2 = pd.read_excel(xls, 'Table2')
table3 = pd.read_excel(xls, 'Table3')

# Perform calculations on each table
# Replace these calculations with your desired operations

# Table 1 calculations
table1_sum = table1.sum()  # Example: calculating sum of columns
table1_mean = table1.mean()  # Example: calculating mean of columns

# Table 2 calculations
table2_sum = table2.sum()  # Example: calculating sum of columns
table2_mean = table2.mean()  # Example: calculating mean of columns

# Table 3 calculations
table3_sum = table3.sum()  # Example: calculating sum of columns
table3_mean = table3.mean()  # Example: calculating mean of columns

# Print the calculated results
print("Table 1 Sum:")
print(table1_sum)
print("\nTable 1 Mean:")
print(table1_mean)

print("\nTable 2 Sum:")
print(table2_sum)
print("\nTable 2 Mean:")
print(table2_mean)

print("\nTable 3 Sum:")
print(table3_sum)
print("\nTable 3 Mean:")
print(table3_mean)
