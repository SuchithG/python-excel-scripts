import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def load_data(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()  # Strip whitespace from column names
    print(f"Loaded sheet: {sheet_name} with columns: {df.columns.tolist()}")
    return df

def calculate_closed_count(df):
    return df['COUNT(*)'].sum()

def calculate_open_assign_count(df, next_month_date, sheet_name):
    # Determine the count column name
    if 'COUNT(*)' in df.columns:
        count_column = 'COUNT(*)'
    elif "COUNT('*')" in df.columns:
        count_column = "COUNT('*')"
    else:
        raise ValueError(f"Count column not found in the sheet {sheet_name}")

    # Check for the presence of 'NOTFCN_ID' and 'NOTFCN_STAT_TYP' columns
    if 'NOTFCN_STAT_TYP' in df.columns and 'NOTFCN_ID' in df.columns:
        # Filter for 'OPEN' and 'CLOSED' statuses
        filtered_df = df[df['NOTFCN_STAT_TYP'].isin(['OPEN', 'CLOSED'])]

        # If you need to consider 'CLOSED' status for a specific date, filter those as well
        # This assumes 'TRUNC(LST_NOTFCN_TMS)' is the correct column name for filtering by date
        if 'TRUNC(LST_NOTFCN_TMS)' in df.columns:
            filtered_df = filtered_df[
                (filtered_df['NOTFCN_STAT_TYP'] != 'CLOSED') |
                (pd.to_datetime(filtered_df['TRUNC(LST_NOTFCN_TMS)']).dt.to_period('M') == next_month_date.to_period('M'))
            ]
        
        # Sum the count column for each unique 'NOTFCN_ID'
        total_count = filtered_df.groupby('NOTFCN_ID')[count_column].sum().reset_index()[count_column].sum()
        return total_count
    else:
        missing_columns = []
        if 'NOTFCN_STAT_TYP' not in df.columns:
            missing_columns.append('NOTFCN_STAT_TYP')
        if 'NOTFCN_ID' not in df.columns:
            missing_columns.append('NOTFCN_ID')
        raise ValueError(f"Required columns {missing_columns} not found in sheet {sheet_name}")


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

open_assign_counts = {}
for loan_type, sheets in sheet_names_open_assign.items():
    total_count = 0
    for sheet in sheets:
        df = load_data(file_path, sheet)
        total_count += calculate_open_assign_count(df, next_month_date)
    open_assign_counts[loan_type] = total_count

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
msg.attach(MIMEText(html_table, 'html'))

# Send the email
with smtplib.SMTP(smtp_host, smtp_port) as server:
    server.starttls() 
    server.login(username, password)
    server.send_message(msg)

print('Email sent!')
