import pandas as pd
import os
from datetime import datetime, timedelta
import random
import smtplib
from email.message import EmailMessage
from email.utils import formataddr

# Function to get dynamic file paths
def get_dynamic_paths(base_path):
    today = datetime.today()
    year = today.strftime('%Y')
    month = today.strftime('%B')
    day = today.strftime('%d %B')
    previous_day = (today - timedelta(days=1)).strftime('%d-%b-%y')
    
    csv_path = os.path.join(base_path, f'FI Exception - {year}', f'{year}', month, f'{day}.csv')
    attendance_tracker_path = os.path.join(base_path, f'FI Exception - {year}', f'{year}', month, 'Attendence Tracker.xlsx')
    assignments_path = os.path.join(base_path, f'FI Exception - {year}', f'{year}', month, 'weekly_assignments.csv')
    return csv_path, attendance_tracker_path, assignments_path, year, previous_day

# Function to send email
def send_email(subject, body, to_emails, attachment_path):
    from_email = 'your_email@example.com'
    from_name = 'Your Name'
    smtp_server = 'smtp.example.com'
    smtp_port = 587
    smtp_user = 'your_email@example.com'
    smtp_password = 'your_password'
    
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = formataddr((from_name, from_email))
    msg['To'] = ', '.join(to_emails)
    msg.set_content(body)
    
    with open(attachment_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)
        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
    
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
    
    print(f"Email sent to {', '.join(to_emails)}")

# Function to load previous assignments
def load_previous_assignments(assignments_path):
    if os.path.exists(assignments_path):
        previous_assignments = pd.read_csv(assignments_path)
        previous_assignments.set_index('Group', inplace=True)
        previous_week = previous_assignments.loc['Week', 'Analyst']
        previous_assignments.drop('Week', inplace=True)
        return previous_assignments, int(previous_week)
    else:
        return pd.DataFrame(columns=['Group', 'Analyst']).set_index('Group'), None

# Function to save current assignments
def save_current_assignments(analyst_mapping, assignments_path):
    df = pd.DataFrame(list(analyst_mapping.items()), columns=['Group', 'Analyst'])
    current_week = get_week_number()
    df = df.append({'Group': 'Week', 'Analyst': current_week}, ignore_index=True)
    df.to_csv(assignments_path, index=False)

# Function to get current week number
def get_week_number():
    return datetime.today().isocalendar()[1]

# Define groups
groups = {
    'Group 1': [16, 587, 70010, 70092],
    'Group 2': [44, 153, 178, 181, 359, 529, 60202, 70006, 70013, 70043, 70093, 90093],
    'Group 3': [3, 173, 206, 224, 591, 70004, 70021],
    'Group 4': [77, 215, 527, 70000, 70015],
    'Group 5': [2, 154, 179, 275, 1001, 90001],
    'Group 6': [15, 23, 75, 588, 70019, 90002],
    'Group 7': [4, 21, 47, 89, 188, 207, 222, 225, 230, 274, 1004, 70025],
    'Group 8': [8, 541, 70005, 70018, 70024],
}

# Base path for the files
base_path = r'G:\'

csv_path, attendance_tracker_path, assignments_path, year, previous_day = get_dynamic_paths(base_path)

# Load the CSV data
csv_data = pd.read_csv(csv_path)

# Load the attendance tracker
attendance_tracker = pd.read_excel(attendance_tracker_path, sheet_name='Sheet1')

# Filter out analysts who are not on leave
available_analysts = attendance_tracker[(attendance_tracker['Leave'] == 'No') & (attendance_tracker['Name'] != 'Karthik')]['Name'].tolist()
unavailable_analysts = attendance_tracker[(attendance_tracker['Leave'] == 'Yes') & (attendance_tracker['Name'] != 'Karthik')]['Name'].tolist()

print(f"Available analysts: {available_analysts}")
print(f"Unavailable analysts: {unavailable_analysts}")

# Calculate Open Exceptions for the current month
attendance_tracker['NOTFCN_CRTE_TMS'] = pd.to_datetime(attendance_tracker['NOTFCN_CRTE_TMS'])
open_exceptions = attendance_tracker[pd.to_datetime(attendance_tracker['NOTFCN_CRTE_TMS']).dt.strftime('%d-%b-%y') == previous_day]['NOTFCN_ID'].nunique()

print(f"Open exceptions for the current month: {open_exceptions}")

# Load previous assignments and get the current week number
previous_assignments, previous_week = load_previous_assignments(assignments_path)
current_week = get_week_number()

analyst_mapping = {}

if previous_week is None or current_week != previous_week:
    # Shuffle the initial list of available analysts to ensure random distribution
    initial_analysts = available_analysts.copy()
    random.shuffle(initial_analysts)
    for i, group in enumerate(groups.keys()):
        analyst_mapping[group] = initial_analysts[i % len(initial_analysts)]
else:
    analyst_mapping = previous_assignments.to_dict()['Analyst']

print(f"New assignments for the week: {analyst_mapping}, Week: {current_week}")

# Reassign groups for today if needed
report_data = []

for group, notifications in groups.items():
    assigned_analyst = analyst_mapping[group]
    if assigned_analyst not in available_analysts:
        available_copy = available_analysts.copy()
        random.shuffle(available_copy)
        temporary_analyst = available_copy.pop()
        analyst_mapping[group] = temporary_analyst
        print(f"Assigned analyst {assigned_analyst} for group {group} is not available. Reassigning to {temporary_analyst} for today.")
    else:
        temporary_analyst = assigned_analyst
    for notification in notifications:
        report_data.append({
            'Notification': notification,
            'Analyst': temporary_analyst,
            'Open Exceptions(for current month)': open_exceptions,
            'Todays Open Exception': csv_data['Todays Open Exception'].sum(),
            'Total Exceptions': csv_data['Total Exceptions'].sum(),
        })

report_df = pd.DataFrame(report_data)

# Save the report to an Excel file
report_file = f'notification_report_{year}.xlsx'
report_df.to_excel(report_file, index=False)

print(f"Report has been generated and saved as {report_file}")

# Save current assignments
save_current_assignments(analyst_mapping, assignments_path)

print(f"Current assignments saved: {analyst_mapping}")

# Send the email with the report
subject = 'Daily Exception Report'
body = 'Please find attached the daily exception report.'
to_emails = ['recipient1@example.com', 'recipient2@example.com']

send_email(subject, body, to_emails, report_file)
