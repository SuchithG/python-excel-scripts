import pandas as pd
import os
import random
from datetime import datetime, timedelta
import smtplib
from email.message import EmailMessage
from email.utils import formataddr

# Function to get dynamic file paths
def get_dynamic_paths(base_path):
    today = datetime.today().strftime("%d-%b-%y")
    previous_day = (datetime.today() - timedelta(days=1)).strftime("%d-%b-%y")

    csv_path = os.path.join(base_path, '28 June.xlsx')
    attendance_tracker_path = os.path.join(base_path, 'Attendance Tracker.xlsx')
    assignments_path = os.path.join(base_path, 'weekly_assignments.xlsx')
    analyst_weeks_path = os.path.join(base_path, 'analyst_weeks.xlsx')

    return csv_path, attendance_tracker_path, assignments_path, analyst_weeks_path, today, previous_day

# Function to load analysts week data
def load_analysts_week(analyst_weeks_path):
    if os.path.exists(analyst_weeks_path):
        df = pd.read_excel(analyst_weeks_path)
        print(f"Loaded analysts week data: \n{df}")
        return df
    else:
        return pd.DataFrame()

# Function to load previous assignments
def load_previous_assignments(assignments_path):
    if os.path.exists(assignments_path):
        df = pd.read_excel(assignments_path, sheet_name='Assignments').set_index('Group')
        print(f"Loaded previous assignments: \n{df}")
        return df, df.loc['Week', 'Analyst']
    else:
        return pd.DataFrame(), None

# Function to save current assignments
def save_current_assignments(analyst_mapping, assignments_path):
    df = pd.DataFrame(list(analyst_mapping.items()), columns=['Group', 'Analyst'])
    current_week = get_week_number()
    df = pd.concat([df, pd.DataFrame([{'Group': 'Week', 'Analyst': current_week}])], ignore_index=True)
    df.to_excel(assignments_path, sheet_name='Assignments', index=False)

# Function to save analysts week data
def save_analysts_week(analysts_week_df, analyst_weeks_path):
    analysts_week_df.to_excel(analyst_weeks_path, index=False)

# Function to get current week number
def get_week_number():
    return datetime.today().isocalendar()[1]

# Define groups
groups = {
    'Group 1': [16, 587, 70010, 700921],
    'Group 2': [255, 600, 70003, 700922],
    'Group 3': [300, 620, 70006, 700923],
    'Group 4': [350, 640, 70008, 700924],
    'Group 5': [400, 660, 70012, 700925],
    'Group 6': [450, 680, 70015, 700926],
    'Group 7': [500, 700, 70018, 700927],
    'Group 8': [550, 720, 70021, 700928]
}

# Load paths
base_path = 'your/base/path/here'
csv_path, attendance_tracker_path, assignments_path, analyst_weeks_path, today, previous_day = get_dynamic_paths(base_path)

# Load the CSV data
csv_data = pd.read_excel(csv_path, sheet_name='28 June')

# Load the attendance tracker
attendance_tracker = pd.read_excel(attendance_tracker_path, sheet_name='Sheet1')

# Filter out analysts who are not on leave
available_analysts = attendance_tracker[(attendance_tracker['Leave'] == 'No') & (attendance_tracker['Name'] != 'Karthik')]['Name'].tolist()
unavailable_analysts = attendance_tracker[(attendance_tracker['Leave'] == 'Yes')]['Name'].tolist()

# Print available and unavailable analysts
print(f"Available analysts: {available_analysts}")
print(f"Unavailable analysts: {unavailable_analysts}")

# Load analysts week data
analysts_week_df = load_analysts_week(analyst_weeks_path)

# If analysts week data is empty, initialize it
if analysts_week_df.empty:
    all_analysts = attendance_tracker[attendance_tracker['Name'] != 'Karthik']['Name'].tolist()
    analysts_week_df = pd.DataFrame({
        'Analyst': all_analysts * (len(groups) // len(all_analysts)) + all_analysts[:len(groups) % len(all_analysts)],
        'Week': list(range(1, len(groups) + 1))
    })
    save_analysts_week(analysts_week_df, analyst_weeks_path)

# Calculate Open Exceptions for the current month
open_exceptions = csv_data[pd.to_datetime(csv_data['NOTFCN_CRTE_TMS']).dt.strftime('%b %Y') == datetime.today().strftime('%b %Y')]['NOTFCN_ID'].nunique()

# Load previous assignments
previous_assignments, previous_week = load_previous_assignments(assignments_path)

# Print debug information
print(f"Previous Assignments: \n{previous_assignments}")
print(f"Previous Week: {previous_week}")

# Assign groups to analysts
analyst_mapping = {}
current_week = get_week_number()
print(f"Current Week: {current_week}")

if previous_week == current_week:
    analyst_mapping = previous_assignments['Analyst'].to_dict()
else:
    week_analysts = analysts_week_df[analysts_week_df['Week'] == current_week % len(groups) + 1]['Analyst'].tolist()
    for i, group in enumerate(groups.keys()):
        analyst_mapping[group] = week_analysts[i % len(week_analysts)]

# Reassign groups if the assigned analyst is not available
for group, notifications in groups.items():
    assigned_analyst = analyst_mapping[group]
    if assigned_analyst not in available_analysts:
        available_copy = available_analysts.copy()
        random.shuffle(available_copy)
        temporary_analyst = available_copy.pop()
        analyst_mapping[group] = temporary_analyst
        print(f"Group {group} reassigned to {temporary_analyst} for today")

# Print final assignments for verification
for group, analyst in analyst_mapping.items():
    print(f"{group} assigned to {analyst}")

# Save the assignments only if it's a new week
if previous_week != current_week:
    save_current_assignments(analyst_mapping, assignments_path)

# Generate the report data
report_data = []
for group, notifications in groups.items():
    assigned_analyst = analyst_mapping[group]
    for notification in notifications:
        report_data.append({
            'Notification': notification,
            'Analyst': assigned_analyst
        })

report_df = pd.DataFrame(report_data)

# Save the report
report_file = f"notification_report_{datetime.today().year}.xlsx"
report_df.to_excel(report_file, index=False)
print(f"Report has been generated and saved as {report_file}")

# Define the send_email function
def send_email(subject, body, to_emails, attachment_path):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = formataddr(('Sender Name', 'sender@example.com'))
    msg['To'] = ', '.join(to_emails)
    msg.set_content(body)

    with open(attachment_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)
        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    with smtplib.SMTP('smtp.example.com', 25) as server:
        server.send_message(msg)
    print(f"Email sent to {', '.join(to_emails)}")

# Send the email with the report
send_email(
    subject="Daily Exception Report",
    body="Please find attached the daily exception report.",
    to_emails=['suchith.girishkumar@db.com'],
    attachment_path=report_file
)
