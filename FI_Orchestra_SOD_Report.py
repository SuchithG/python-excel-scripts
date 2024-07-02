import pandas as pd
import os
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from openpyxl import load_workbook
import random
from datetime import datetime, timedelta

# Function to get dynamic file paths
def get_dynamic_paths(base_path):
    today = datetime.today()
    year = today.strftime('%Y')
    month = today.strftime('%B')
    day = today.strftime('%d %B')
    previous_day = (today - timedelta(days=1)).strftime('%d-%b-%y')
    
    csv_path = os.path.join(base_path, f'FI Exception - {year}', f'{year}', month, f'{day}.xlsx')
    attendance_tracker_path = os.path.join(base_path, f'FI Exception - {year}', f'{year}', month, 'Attendence Tracker.xlsx')
    assignments_path = os.path.join(base_path, f'FI Exception - {year}', 'weekly_assignments.csv')
    
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

# Function to load previous assignments
def load_previous_assignments(assignments_path):
    if os.path.exists(assignments_path):
        return pd.read_csv(assignments_path).set_index('Group').to_dict()['Analyst']
    else:
        return {}

# Function to save current assignments
def save_current_assignments(analyst_mapping, assignments_path):
    df = pd.DataFrame(list(analyst_mapping.items()), columns=['Group', 'Analyst'])
    df.to_csv(assignments_path, index=False)

# Function to get current week number
def get_week_number():
    return datetime.today().isocalendar()[1]

# Function to check if it's a weekday (Monday to Friday)
def is_weekday():
    return datetime.today().weekday() < 5

# Base path for the files
base_path = r'/mnt/data/'

csv_path, attendance_tracker_path, assignments_path, year, previous_day = get_dynamic_paths(base_path)

# Check if today is a weekday
if is_weekday():
    # Load the data from the Excel file with specified engine
    workbook = load_workbook(filename=csv_path, data_only=True)
    sheet = workbook['28 June']

    # Convert the worksheet to a DataFrame
    data = pd.DataFrame(sheet.values)

    # Assign the first row as column names
    data.columns = data.iloc[0]
    data = data[1:]

    # Load the attendance tracker
    attendance_tracker = pd.read_excel(attendance_tracker_path, sheet_name='Sheet1')

    # Filter out "Karthik" and analysts who are not on leave
    available_analysts = attendance_tracker[(attendance_tracker['Leave'] == 'No') & (attendance_tracker['Name'] != 'Karthik')]['Name'].tolist()
    unavailable_analysts = attendance_tracker[(attendance_tracker['Leave'] == 'Yes')]['Name'].tolist()

    # Debug print statements
    print("Available analysts: ", available_analysts)
    print("Unavailable analysts: ", unavailable_analysts)

    # Calculate Open Exceptions for the current month
    data['NOTFCN_CRTE_TMS'] = pd.to_datetime(data['NOTFCN_CRTE_TMS'])
    open_exceptions = data[data['NOTFCN_CRTE_TMS'].dt.strftime('%d-%b-%y') == previous_day]['NOTFCN_ID'].nunique()

    # Debug print statement
    print("Open exceptions for the current month: ", open_exceptions)

    # Load previous assignments if they exist
    previous_assignments = load_previous_assignments(assignments_path)

    # Check if it's a new week
    current_week = get_week_number()
    if previous_assignments and previous_assignments.get('Week') == current_week:
        analyst_mapping = {k: v for k, v in previous_assignments.items() if k != 'Week'}
        print("Using previous assignments for the week:", analyst_mapping)
    else:
        # Create new weekly assignments
        analyst_mapping = {}
        initial_analysts = available_analysts.copy()
        random.shuffle(initial_analysts)  # Shuffle the list to ensure random distribution
        for i, group in enumerate(groups.keys()):
            analyst_mapping[group] = initial_analysts[i % len(initial_analysts)]
        # Add current week to the mapping for reference
        analyst_mapping['Week'] = current_week
        # Save current assignments
        save_current_assignments(analyst_mapping, assignments_path)
        print("New assignments for the week:", analyst_mapping)

    # Assign analysts to notifications, ensuring that unavailable analysts are handled
    report_data = []

    for group, notifications in groups.items():
        assigned_analyst = analyst_mapping.get(group, 'No Analyst Assigned')
        if assigned_analyst in unavailable_analysts:
            print(f"Assigned analyst {assigned_analyst} for group {group} is not available. Reassigning for today...")
            available_copy = available_analysts.copy()
            random.shuffle(available_copy)  # Shuffle the list to ensure randomness
            temporary_analyst = available_copy.pop() if available_copy else 'No Analyst Available'
            print(f"Group {group} reassigned to {temporary_analyst} for today")
            for notification in notifications:
                report_data.append({
                    'Notification': notification,
                    'Analyst': temporary_analyst,
                    'Open Exceptions(for current month)': open_exceptions,
                    'Todays Open Exception': data['Todays Open Exception'].sum() if 'Todays Open Exception' in data.columns else 0,
                    'Total Exceptions': data['Total Exceptions'].sum() if 'Total Exceptions' in data.columns else 0,
                })
        else:
            for notification in notifications:
                report_data.append({
                    'Notification': notification,
                    'Analyst': assigned_analyst,
                    'Open Exceptions(for current month)': open_exceptions,
                    'Todays Open Exception': data['Todays Open Exception'].sum() if 'Todays Open Exception' in data.columns else 0,
                    'Total Exceptions': data['Total Exceptions'].sum() if 'Total Exceptions' in data.columns else 0,
                })
            print(f"Group {group} assigned to {assigned_analyst}")

    report_df = pd.DataFrame(report_data)

    # Save the report to an Excel file
    report_file = f'notification_report_{year}.xlsx'
    report_df.to_excel(report_file, index=False)

    print(f"Report has been generated and saved as {report_file}")

    # Send the email with the report
    subject = 'Daily Exception Report'
    body = 'Please find attached the daily exception report.'
    to_emails = ['suchith.girishkumar@db.com']

    send_email(subject, body, to_emails, report_file)
else:
    print("Today is a weekend. The script will not run.")
