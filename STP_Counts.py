import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def load_data(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()  # Strip whitespace from column names
    return df

def calculate_closed_count(df):
    count_column = 'COUNT(*)' if 'COUNT(*)' in df.columns else "COUNT('*')"
    return df[count_column].sum()

def calculate_combined_unique_open_assign_count(file_path, sheets, next_month_date):
    combined_df = pd.DataFrame()
    for sheet in sheets:
        df = load_data(file_path, sheet)
        combined_df = pd.concat([combined_df, df], ignore_index=True)

    # Determine the correct count column name
    count_column = 'COUNT(*)' if 'COUNT(*)' in combined_df.columns else "COUNT('*')"

    # Convert 'TRUNC(LST_NOTFCN_TMS)' to datetime
    combined_df['TRUNC(LST_NOTFCN_TMS)'] = pd.to_datetime(combined_df['TRUNC(LST_NOTFCN_TMS)'], errors='coerce')

    # Apply filters
    filtered_combined_df = combined_df[
        (combined_df['NOTFCN_STAT_TYP'] == 'OPEN') |
        ((combined_df['NOTFCN_STAT_TYP'] == 'CLOSED') &
         (combined_df['TRUNC(LST_NOTFCN_TMS)'].dt.month == next_month_date.month) &
         (combined_df['TRUNC(LST_NOTFCN_TMS)'].dt.year == next_month_date.year))
    ]

    # Sum the count for unique NOTFCN_IDs
    unique_counts = filtered_combined_df.drop_duplicates(subset='NOTFCN_ID')[count_column].sum()
    return unique_counts

def calculate_ageing_breaks(df, next_month_date):
    # Determine the correct count column name
    count_column = 'COUNT(*)' if 'COUNT(*)' in df.columns else "COUNT('*')"
    
    # Calculate the age of each item based on 'TRUNC(NOTFCN_CRTE_TMS)' or 'TRUNC(LST_NOTFCN_TMS)'
    df['Age Open'] = (next_month_date - pd.to_datetime(df['TRUNC(NOTFCN_CRTE_TMS)'])).dt.days
    df['Age Closed'] = (next_month_date - pd.to_datetime(df['TRUNC(LST_NOTFCN_TMS)'])).dt.days

    # Define the bins for the ageing categories
    bins = [-1, 1, 7, 15, 30, 180, float('inf')]
    labels = ['0-1 New', '02-07 days', '08-15 days', '16-30 days', '31-180 days', '>180 days']
    
    # Initialize the counts for each category
    ageing_breaks = {label: 0 for label in labels}

    # Categorize open and closed items
    for _, row in df.iterrows():
        if row['NOTFCN_STAT_TYP'] == 'OPEN':
            age = row['Age Open']
        elif row['NOTFCN_STAT_TYP'] == 'CLOSED' and pd.to_datetime(row['TRUNC(LST_NOTFCN_TMS)']).month == next_month_date.month:
            age = row['Age Closed']
        else:
            continue

        # Increment the correct ageing category
        for label, upper_bound in zip(labels, bins[1:]):
            if age < upper_bound:
                ageing_breaks[label] += row[count_column]
                break

    return ageing_breaks

# Replace with your actual file path
file_path = 'your_file.xlsx'

# Closed counts
sheets_and_loans_closed = {
    'Line 180': 'Loans',
    'Line 1280': 'FI',
    'Line 655': 'Equity',
    'Line 2020': 'LD'
}

closed_counts = {}
for sheet, loan_type in sheets_and_loans_closed.items():
    df = load_data(file_path, sheet)
    closed_count = calculate_closed_count(df)
    closed_counts[loan_type] = closed_count

# Open/Assign counts
next_month_date = pd.Timestamp('2023-12-01')  # Set this to the next month date you're interested in

sheet_names_open_assign = {
    'Loans': ['Line 270', 'Line 297', 'Line 441', 'Line 523'],
    'FI': ['Line 1616', 'Line 1407', 'Line 1727', 'Line 1843'],
    'Equity': ['Line 764', 'Line 809', 'Line 970', 'Line 1024', 'Line 1088'],
    'LD': ['Line 2104', 'Line 2261', 'Line 2325', 'Line 2389']
}

'''
open_assign_counts = {}
for loan_type, sheets in sheet_names_open_assign.items():
    total_count = 0
    for sheet in sheets:
        df = load_data(file_path, sheet)
        total_count += calculate_open_assign_count(df, next_month_date, sheet)
    open_assign_counts[loan_type] = total_count
'''
# Function to convert the open ageing breaks data to a DataFrame and then to HTML
def format_open_ageing_breaks_to_html(ageing_breaks):
    # Convert the dictionary to a DataFrame
    ageing_breaks_df = pd.DataFrame(ageing_breaks).T
    ageing_breaks_df.columns = ['0-1 New', '02-07 days', '08-15 days', '16-30 days', '31-180 days', '>180 days']
    ageing_breaks_df.index.name = 'Loan Type'

    # Convert the DataFrame to HTML
    return ageing_breaks_df.to_html()

# Calculate closed and open/assign counts
closed_counts = {loan_type: calculate_closed_count(load_data(file_path, sheet)) for sheet, loan_type in sheets_and_loans_closed.items()}
open_assign_counts = {loan_type: calculate_combined_unique_open_assign_count(file_path, sheets, next_month_date) for loan_type, sheets in sheet_names_open_assign.items()}

# Calculate open ageing breaks for each loan type
open_ageing_breaks = {}
for loan_type, sheets in sheet_names_open_assign.items():
    combined_df = pd.concat([load_data(file_path, sheet) for sheet in sheets], ignore_index=True)
    open_ageing_breaks[loan_type] = calculate_ageing_breaks(combined_df, next_month_date)

# Convert the open ageing breaks to HTML
open_ageing_breaks_html = format_open_ageing_breaks_to_html(open_ageing_breaks)


# Creating a DataFrame for email content
data_for_email = {
    'Loan Type': [],
    'Closed Count': [],
    'Open/Assign Count': []
}
for loan_type in closed_counts:
    data_for_email['Loan Type'].append(loan_type)
    data_for_email['Closed Count'].append(closed_counts[loan_type])
    data_for_email['Open/Assign Count'].append(open_assign_counts.get(loan_type, 0))

email_df = pd.DataFrame(data_for_email)
html_table = email_df.to_html(index=False)

# Combine both tables' HTML content
combined_html_table = html_table + "<br><br>" + open_ageing_breaks_html

# Email setup (replace with your actual details)
smtp_host = 'your_smtp_host'
smtp_port = your_smtp_port
username = 'your_username'
password = 'your_password'
sender_email = 'sender@example.com'
recipient_email = 'recipient@example.com'

# Email content
msg = MIMEMultipart()
msg['Subject'] = 'Loan Counts Table'
msg['From'] = sender_email
msg['To'] = recipient_email
msg.attach(MIMEText(combined_html_table, 'html'))

# Send the email
with smtplib.SMTP(smtp_host, smtp_port) as server:
    server.starttls() 
    server.login(username, password)
    server.send_message(msg)

print('Email sent!')
