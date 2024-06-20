import pandas as pd
import os
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Function to log dataframes for debugging
def log_dataframe(df, name):
    print(f"\n{name} DataFrame:")
    print(df.head())

# Get the current year and month
current_year = datetime.now().year
current_month = datetime.now().strftime('%B %Y')
current_day = datetime.now().strftime('%d %B')

# Paths to uploaded files
input_file_path = '/mnt/data/file-H3v1VwRRi36DE7ml4zb1wb8P'  

# Read the Excel file for the first two tables
df_leave_tracker = pd.read_excel(input_file_path, sheet_name='Leave Tracker')
df_daily_report = pd.read_excel(input_file_path, sheet_name='18 June')  

# Log dataframes
log_dataframe(df_leave_tracker, "Leave Tracker")
log_dataframe(df_daily_report, "Daily Report")

# Read the Excel file for the third and fourth tables
third_fourth_table_path = 'G:/Sreekanth/Jun 2024/19-Jun-2024/FI MAIN HUB REPORT 2024.xlsx'

# Third table
df_third_table = pd.read_excel(third_fourth_table_path, sheet_name='INC STATUS')

# Fourth table
df_fourth_table = pd.read_excel(third_fourth_table_path, sheet_name='VENDOR TICKETS')

# Log dataframes
log_dataframe(df_third_table, "Third Table")
log_dataframe(df_fourth_table, "Fourth Table")

# Process the Leave Tracker data to include only "Name", "Shift", and "Attendance"
df_leave_tracker_processed = df_leave_tracker[['Name', 'Shift', 'Attendance']]

# Process the Daily Report data
df_daily_report_processed = df_daily_report[[
    'Notifications Id', 'Analyst', 'Open Exceptions (for current month)', 
    "Today's Open Exception", 'Total Exceptions', 
    "Count of exceptions closed today's", 'Known issues', 
    'Exception Count WIP (with vendor/ L2)', 'Exceptions count EOD', 
    'Dependency', 'Comments'
]]

# Replace NaN with blank and explicitly convert float columns to int where necessary
df_daily_report_processed = df_daily_report_processed.fillna('')
float_columns = df_daily_report_processed.select_dtypes(include=['float64']).columns

# Log before conversion
print("\nBefore conversion to integers:")
print(df_daily_report_processed[float_columns].head())

# Convert float columns to integers where appropriate
for col in float_columns:
    df_daily_report_processed[col] = df_daily_report_processed[col].apply(lambda x: int(x) if pd.notnull(x) and x.is_integer() else x)

# Convert all numbers to string format
for col in df_daily_report_processed.columns:
    if df_daily_report_processed[col].dtype in ['float64', 'int64']:
        df_daily_report_processed[col] = df_daily_report_processed[col].apply(lambda x: '{:.0f}'.format(x) if pd.notnull(x) else '')

# Log after conversion
print("\nAfter conversion to integers:")
print(df_daily_report_processed[float_columns].head())

# Filter the third table data
today_str = datetime.now().strftime('%d-%B-%Y')
df_third_table_filtered = df_third_table[(df_third_table['Ticket Status'] == 'Open') & 
                                         (df_third_table['Latest update from L2/L3/Ops'].str.contains(today_str))]

# Select the required columns for the third table
df_third_table_processed = df_third_table_filtered[[
    'Notification ID', 'Raised By', 'DBUnity Incident #', 'Incident Subject', 'Created Date', 
    'Status', 'Assignee', 'Latest update from L2/L3/Ops', 'Pending with', 'Priority defined by Ops'
]]

# Replace NaN with blank
df_third_table_processed = df_third_table_processed.fillna('')

# Filter the fourth table data
df_fourth_table_filtered = df_fourth_table[(df_fourth_table['Ticket Status'] == 'Open') & 
                                           (df_fourth_table['Comments'].str.contains(today_str))]

# Select the required columns for the fourth table
df_fourth_table_processed = df_fourth_table_filtered[[
    'Raised BY', 'Vendor', 'Vendor ticket No.', 'Created Date', 'Ticket Status', 
    'Closed Date', 'Notification ID', 'BBG call dis', 'Comments', 'Updated Status'
]]

# Replace NaN with blank
df_fourth_table_processed = df_fourth_table_processed.fillna('')

# Convert Leave Tracker DataFrame to HTML with light green header background
leave_tracker_html = df_leave_tracker_processed.to_html(index=False, border=1)
leave_tracker_html = leave_tracker_html.replace('<thead>', '<thead style="background-color: lightgreen;">')

# Function to highlight rows with blank "Comments" in yellow
def highlight_rows(df):
    rows = df.to_dict(orient='records')
    html = "<table border='1' cellspacing='0' cellpadding='5' style='border-collapse: collapse; width: 100%;'>"
    html += "<thead><tr style='background-color: #4CAF50; color: white;'>"
    for col in df.columns:
        html += f"<th>{col}</th>"
    html += "</tr></thead><tbody>"
    for row in rows:
        if row.get('Comments', '') == '':
            html += "<tr style='background-color: yellow'>"
        else:
            html += "<tr>"
        for col in df.columns:
            html += f"<td>{row[col]}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html

# Apply the function to the daily report
styled_daily_report_html = highlight_rows(df_daily_report_processed)

# Convert the third table DataFrame to HTML
third_table_html = df_third_table_processed.to_html(index=False, border=1)
third_table_html = third_table_html.replace('<thead>', '<thead style="background-color: lightblue;">')

# Convert the fourth table DataFrame to HTML
fourth_table_html = df_fourth_table_processed.to_html(index=False, border=1)
fourth_table_html = fourth_table_html.replace('<thead>', '<thead style="background-color: lightyellow;">')

# Debug: Verify the generated HTML
print("\nLeave Tracker HTML:")
print(leave_tracker_html)
print("\nStyled Daily Report HTML:")
print(styled_daily_report_html)
print("\nThird Table HTML:")
print(third_table_html)
print("\nFourth Table HTML:")
print(fourth_table_html)

# Email details
smtp_server = 'smtp.example.com'
smtp_port = 587
smtp_user = 'your_email@example.com'
smtp_password = 'your_password'
from_email = 'your_email@example.com'
to_email = 'recipient@example.com'
subject = 'EOD Report'

# Create the email content
email_content = f"""
<html>
<head>
<style>
    table {{
        border-collapse: collapse;
        width: 100%;
    }}
    th, td {{
        text-align: left;
        padding: 8px;
        border: 1px solid black;
    }}
    th {{
        background-color: lightgreen;
        color: black;
    }}
    .highlight {{
        background-color: yellow;
    }}
</style>
</head>
<body>
    <h2>Leave Tracker</h2>
    {leave_tracker_html}
    <h2>Daily Report ({current_day})</h2>
    {styled_daily_report_html}
    <h2>Incident Status</h2>
    {third_table_html}
    <h2>Vendor Tickets</h2>
    {fourth_table_html}
</body>
</html>
"""

# Create the MIME email
msg = MIMEMultipart('alternative')
msg['Subject'] = subject
msg['From'] = from_email
msg['To'] = to_email
part = MIMEText(email_content, 'html')
msg.attach(part)

# Send the email
with smtplib.SMTP(smtp_server, smtp_port) as server:
    server.starttls()
    server.login(smtp_user, smtp_password)
    server.sendmail(from_email, to_email, msg.as_string())

print("EOD report has been sent to:", to_email)
