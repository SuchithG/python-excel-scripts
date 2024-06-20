import pandas as pd
import os
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Get the current year and month
current_year = datetime.now().year
current_month = datetime.now().strftime('%B %Y')
current_day = datetime.now().strftime('%d %B')

# Paths to uploaded files
input_file_path = '/mnt/data/file-SWqCZUZO8umQV5AtKKsMWDTc'  # Use the relevant file uploaded for the first and second tables

# Read the Excel file for the first two tables
df_leave_tracker = pd.read_excel(input_file_path, sheet_name='Leave Tracker')
df_daily_report = pd.read_excel(input_file_path, sheet_name='18 June')  # Assuming this is the correct sheet name

# Read the Excel file for the third table
third_table_path = 'G:/Sreekanth/Jun 2024/19-Jun-2024/FI MAIN HUB REPORT 2024.xlsx'
df_third_table = pd.read_excel(third_table_path, sheet_name='INC STATUS')

# Debug: Verify data read from Excel
print("Leave Tracker Data:")
print(df_leave_tracker.head())
print("\nDaily Report Data:")
print(df_daily_report.head())
print("\nThird Table Data:")
print(df_third_table.head())

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

# Replace NaN with blank
df_daily_report_processed = df_daily_report_processed.fillna('')

# Convert float to integer where possible
def convert_floats_to_int(df):
    for col in df.select_dtypes(include=['float']):
        df[col] = df[col].apply(lambda x: '{:.0f}'.format(x) if x.is_integer() else x)
    return df

df_daily_report_processed = convert_floats_to_int(df_daily_report_processed)

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
        if row['Comments'] == '':
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

# Debug: Verify the generated HTML
print("\nLeave Tracker HTML:")
print(leave_tracker_html)
print("\nStyled Daily Report HTML:")
print(styled_daily_report_html)
print("\nThird Table HTML:")
print(third_table_html)

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
    <h2>Third Table</h2>
    {third_table_html}
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
