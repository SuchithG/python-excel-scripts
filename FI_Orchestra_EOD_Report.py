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
input_file_path = '/mnt/data/file-cqRszSIpLiCna4ApkrdwQRdR'  # Use the relevant file uploaded

# Read the Excel file
df_leave_tracker = pd.read_excel(input_file_path, sheet_name='Leave Tracker')
df_daily_report = pd.read_excel(input_file_path, sheet_name='18 June')  # Assuming this is the correct sheet name

# Debug: Verify data read from Excel
print("Leave Tracker Data:")
print(df_leave_tracker.head())
print("\nDaily Report Data:")
print(df_daily_report.head())

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
        df[col] = df[col].apply(lambda x: int(x) if x.is_integer() else x)
    return df

df_daily_report_processed = convert_floats_to_int(df_daily_report_processed)

# Convert Leave Tracker DataFrame to HTML with light green header background
leave_tracker_html = df_leave_tracker_processed.to_html(index=False, border=1)

leave_tracker_html = leave_tracker_html.replace(
    '<thead>',
    '<thead style="background-color: lightgreen;">'
)

# Function to highlight rows with blank "Comments" in yellow and set red color for specific columns
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
            if col in ['Total Exceptions', 'Exceptions count EOD']:
                html += f"<td style='color: red'>{row[col]}</td>"
            else:
                html += f"<td>{row[col]}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html

# Apply the function to the daily report
styled_daily_report_html = highlight_rows(df_daily_report_processed)

# Debug: Verify the generated HTML
print("\nLeave Tracker HTML:")
print(leave_tracker_html)
print("\nStyled Daily Report HTML:")
print(styled_daily_report_html)

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
    .red-text {{
        color: red;
    }}
</style>
</head>
<body>
    <h2>Leave Tracker</h2>
    {leave_tracker_html}
    <h2>Daily Report ({current_day})</h2>
    {styled_daily_report_html}
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
