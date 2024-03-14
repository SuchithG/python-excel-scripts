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
bcc_recipients = ['bcc@example.com']  # Optional, for demonstration

# Function to get the three required dates excluding Sundays
def get_required_dates():
    required_dates = []
    current_date = datetime.now()
    while len(required_dates) < 3:
        current_date -= timedelta(days=1)
        if current_date.weekday() != 6:  # Skip Sundays
            required_dates.append(current_date)
    return required_dates

# Function to get the current month as a string in "MMM-YY" format
def get_current_month_string():
    return datetime.now().strftime('%b-%y')

# Function to check for data presence on any of the required dates for the current month
def check_consecutive_dates(file_path, current_month):
    try:
        df = pd.read_excel(file_path)
        # Filter for the current month and 'EQ' Asset Class
        df_filtered = df[(df['Month'] == current_month) & (df['Asset Class'] == 'EQ')].copy()
        df_filtered['Date'] = pd.to_datetime(df_filtered['Date'], format='%d-%b-%y')
        required_dates = get_required_dates()
        # Check if any required date is present in the data
        data_exists = any((df_filtered['Date'].dt.date == date.date()).any() for date in required_dates)
        return not data_exists
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return True

# List to store names of files without data on any of the required dates
missing_data_files = []

# Check each Excel workbook in the folder
current_month = get_current_month_string()
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        if check_consecutive_dates(file_path, current_month):
            missing_data_files.append(file_name)

# Function to send email with the list of files
def send_email(missing_files, to_recipients, cc_recipients):
    body = "The following Excel files have no data under 'EQ' Asset Class for any of the three consecutive working days:\n\n" + "\n".join(missing_files)
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = ', '.join(to_recipients)
    msg['CC'] = ', '.join(cc_recipients)
    msg['Subject'] = 'Missing Data for Consecutive Working Days in Excel Files'
    msg.attach(MIMEText(body, 'plain'))
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
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
    print("All files have data for at least one of the three consecutive working days under the 'EQ' Asset Class.")
