import os
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Define the folder path where your Excel workbooks are stored
folder_path = '/path/to/excel/folder'

# Define SMTP credentials
smtp_server = 'smtp.example.com'
smtp_port = 587  # or 465 for SSL
smtp_user = 'your_email@example.com'
smtp_password = 'your_password'

# Define email recipients
to_recipients = ['recipient1@example.com', 'recipient2@example.com']
cc_recipients = ['cc1@example.com', 'cc2@example.com']
bcc_recipients = ['bcc@example.com']  # For demonstration, you can add BCC recipients if needed

# Function to get the three required dates excluding Sundays
def get_required_dates():
    required_dates = []
    current_date = datetime.now()
    while len(required_dates) < 3:
        current_date -= timedelta(days=1)
        if current_date.weekday() != 6:  # 6 corresponds to Sunday
            required_dates.append(current_date)
    return required_dates

def check_consecutive_dates(file_path):
    try:
        df = pd.read_excel(file_path)
        # Filter rows where Asset Class is 'EQ'
        df_eq = df[df['Asset Class'] == 'EQ']
        # Explicitly specify the date format for parsing
        df_eq['Date'] = pd.to_datetime(df_eq['Date'], format='%Y-%m-%d')
        required_dates = get_required_dates()
        # Check if data exists for at least one of the required dates
        data_exists = any(df_eq['Date'].dt.date == date.date() for date in required_dates)
        return not data_exists  # Return True if data does NOT exist for all required dates
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return True  # Consider files with errors as needing attention

# List to store names of files that do not have data for three consecutive working days
missing_data_files = []

# Check each Excel workbook in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):  # Ensure we're only looking at Excel files
        file_path = os.path.join(folder_path, file_name)
        if check_consecutive_dates(file_path):
            missing_data_files.append(file_name)

# Function to send email
def send_email(missing_files, to_recipients, cc_recipients):
    body = "The following Excel files do not have data for three consecutive working days under the 'EQ' Asset Class:\n\n" + "\n".join(missing_files)
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = ', '.join(to_recipients)
    msg['CC'] = ', '.join(cc_recipients)  # Include CC recipients in the header
    msg['Subject'] = 'Excel Files Missing Consecutive Working Days Data for EQ Asset Class'
    msg.attach(MIMEText(body, 'plain'))
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        # Combine TO, CC, and potentially BCC lists for the recipient parameter in sendmail
        all_recipients = to_recipients + cc_recipients + bcc_recipients
        text = msg.as_string()
        server.sendmail(smtp_user, all_recipients, text)
        server.quit()
        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Send the email if there are files that do not meet the criteria
if missing_data_files:
    send_email(missing_data_files, to_recipients, cc_recipients)
else:
    print("All files have data for three consecutive working days under the 'EQ' Asset Class.")
