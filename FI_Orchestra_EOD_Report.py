import pandas as pd
import os
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime import base

# Get the current year and month
current_year = datetime.now().year
current_month = datetime.now().strftime('%B %Y')
current_day = datetime.now().strftime('%d %B')

# Construct the input file path dynamically
base_dir = r'G:\{}\FI Exception - {}\SOD and EOD Report'.format(current_year, current_year)
input_file_name = '{} {} SOD and EOD Report Updated Version.xlsx'.format(current_month, current_year)
input_file_path = os.path.join(base_dir, input_file_name)

# Read the Excel file
df_leave_tracker = pd.read_excel(input_file_path, sheet_name='Leave Tracker')
df_daily_report = pd.read_excel(input_file_path, sheet_name=current_day)

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

# Apply formatting: highlight rows with blank "Comments" in light yellow
def highlight_blank_comments(row):
    return ['background-color: yellow' if pd.isna(row['Comments']) else '' for _ in row]

# Convert DataFrames to HTML with formatting
leave_tracker_html = df_leave_tracker_processed.to_html(index=False)

# Apply style to the daily report
styled_daily_report = df_daily_report_processed.style.apply(highlight_blank_comments, axis=1)\
                                                     .set_properties(**{'color': 'red'}, subset=['Total Exceptions', 'Exceptions count EOD'])\
                                                     .render()

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
<head></head>
<body>
    <h2>Leave Tracker</h2>
    {leave_tracker_html}
    <h2>Daily Report ({current_day})</h2>
    {styled_daily_report}
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
