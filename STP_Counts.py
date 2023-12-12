import pandas as pd

# Load the data from each sheet into a DataFrame
def load_data(sheet_name):
    df = pd.read_excel('your_file.xlsx', sheet_name=sheet_name)
    return df

# Calculate the sum of the "COUNT(*)" column for the given DataFrame
def calculate_closed_count(df):
    return df['COUNT(*)'].sum()

# Replace 'your_file.xlsx' with the path to your actual Excel file
file_path = 'your_file.xlsx'

# Calculate closed counts for each loan type
sheets_and_loans = {
    'Line 180': 'Loans',
    'Line 1280': 'FI',
    'Line 655': 'Equity',
    'Line 2020': 'LD'
}

closed_counts = {}

for sheet, loan_type in sheets_and_loans.items():
    df = load_data(sheet)
    closed_count = calculate_closed_count(df)
    closed_counts[loan_type] = closed_count

# Create a DataFrame for the closed count data
closed_counts_df = pd.DataFrame(list(closed_counts.items()), columns=['Loan Type', 'Closed Count'])

# Display the DataFrame as a styled HTML table
styled_table = closed_counts_df.style.set_table_styles(
    [{
        'selector': 'th',
        'props': [('background-color', '#FFFF00'), ('color', 'black')]
    }]
).set_properties(**{
    'background-color': 'white',
    'color': 'black',
    'border-color': 'black',
    'border-style' :'solid',
    'border-width': '1px'
}).set_caption("Closed Counts")

# Display styled table
styled_table
