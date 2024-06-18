import pandas as pd
import os
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Get the current year and month
current_year = datetime.now().year
current_month = datetime.now().strftime('%B %Y')

# Construct the input file path dynamically
base_dir = r'G:\{}\FI Exception - {}\SOD and EOD Report'.format(current_year, current_year)
input_file_name = '{} {} SOD and EOD Report Updated Version.xlsx'.format(current_month, current_year)
input_file_path = os.path.join(base_dir, input_file_name)

# Read the Excel file
df_leave_tracker = pd.read_excel(input_file_path, sheet_name='Leave Tracker')

# Process the Leave Tracker data to include only "Name", "Shift", and "Attendance"
df_leave_tracker_processed = df_leave_tracker[['Name', 'Shift', 'Attendance']]

# Convert DataFrame to HTML
leave_tracker_html = df_leave_tracker_processed.to_html(index=False)

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
