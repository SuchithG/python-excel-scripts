import pandas as pd
import os
import random
from datetime import datetime, timedelta
import smtplib
from email.message import EmailMessage
from email.utils import formataddr

# Function to get dynamic file paths
def get_dynamic_paths(base_path):
    analyst_weeks_path = os.path.join(base_path, 'analyst_weeks.xlsx')
    weekly_assignments_path = os.path.join(base_path, 'weekly_assignments.xlsx')
    attendance_tracker_path = os.path.join(base_path, 'Attendance Tracker.xlsx')
    notifications_data_path = os.path.join(base_path, '28 June.xlsx')

    return analyst_weeks_path, weekly_assignments_path, attendance_tracker_path, notifications_data_path

# Function to get current week number
def get_week_number():
    return datetime.strptime('28-Jun-24', '%d-%b-%y').isocalendar()[1]

# Function to load analysts week data
def load_analysts_week(analyst_weeks_path):
    if os.path.exists(analyst_weeks_path):
        df = pd.read_excel(analyst_weeks_path)
        print(f"Loaded analysts week data: \n{df}")
        return df
    else:
        return pd.DataFrame()

# Function to save analysts week data
def save_analysts_week(analysts_week_df, analyst_weeks_path):
    analysts_week_df.to_excel(analyst_weeks_path, index=False)

# Function to load weekly assignments
def load_weekly_assignments(weekly_assignments_path):
    if os.path.exists(weekly_assignments_path):
        df = pd.read_excel(weekly_assignments_path)
        print(f"Loaded weekly assignments: \n{df}")
        return df
    else:
        return pd.DataFrame()

# Function to save weekly assignments
def save_weekly_assignments(weekly_assignments_df, weekly_assignments_path):
    weekly_assignments_df.to_excel(weekly_assignments_path, index=False)

# Function to assign notifications
def assign_notifications(analyst_mapping, available_analysts, groups):
    assignments = {analyst: [] for analyst in available_analysts}
    unassigned_notifications = []

    for group, analyst in analyst_mapping.items():
        if analyst in available_analysts:
            assignments[analyst].extend(groups[group])
        else:
            unassigned_notifications.extend(groups[group])

    # Distribute unassigned notifications among available analysts
    while unassigned_notifications:
        for analyst in available_analysts:
            if unassigned_notifications:
                assignments[analyst].append(unassigned_notifications.pop(0))

    return assignments

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
output_path = 'your/output/path/here'  # Specify the folder where you want to save the output file
analyst_weeks_path, weekly_assignments_path, attendance_tracker_path, notifications_data_path = get_dynamic_paths(base_path)

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

# Get current week
current_week = get_week_number()

# If analysts week data is empty or new week, initialize it
if analysts_week_df.empty or current_week not in analysts_week_df['Week'].values:
    all_analysts = attendance_tracker[attendance_tracker['Name'] != 'Karthik']['Name'].tolist()
    previous_week = analysts_week_df['Week'].max() if not analysts_week_df.empty else current_week - 1
    previous_week_analysts = analysts_week_df[analysts_week_df['Week'] == previous_week].sort_values('Group')['Analyst'].tolist()
    new_week_analysts = pd.DataFrame({
        'Analyst': previous_week_analysts[1:] + [previous_week_analysts[0]],  # Round-robin rotation
        'Group': list(groups.keys()),
        'Week': current_week
    })
    analysts_week_df = pd.concat([analysts_week_df, new_week_analysts], ignore_index=True)
    save_analysts_week(analysts_week_df, analyst_weeks_path)

# Get the current week's analyst mapping
current_week_mapping = analysts_week_df[analysts_week_df['Week'] == current_week].set_index('Group')['Analyst'].to_dict()

# Assign notifications
assignments = assign_notifications(current_week_mapping, available_analysts, groups)

# Prepare weekly assignments dataframe
weekly_assignments_data = []
for analyst, notifications in assignments.items():
    for notification in notifications:
        weekly_assignments_data.append({
            'Analyst': analyst,
            'Notification': notification
        })

weekly_assignments_df = pd.DataFrame(weekly_assignments_data)
save_weekly_assignments(weekly_assignments_df, weekly_assignments_path)

# Print final assignments for verification
print(f"Weekly Assignments: \n{weekly_assignments_df}")

# Load the notifications data
notifications_data = pd.read_excel(notifications_data_path, sheet_name='28 June')

# Hardcode today's date
hardcoded_date = '28-Jun-24'
hardcoded_date_dt = datetime.strptime(hardcoded_date, '%d-%b-%y').date()

# Calculate Open Exceptions for the current month excluding hardcoded date
open_exceptions_current_month = notifications_data[
    (pd.to_datetime(notifications_data['NOTFCN_CRTE_TMS']).dt.strftime('%b %Y') == hardcoded_date_dt.strftime('%b %Y')) &
    (pd.to_datetime(notifications_data['NOTFCN_CRTE_TMS']).dt.date != hardcoded_date_dt)
]['NOTFCN_ID'].value_counts().to_dict()

# Calculate Today's Open Exceptions
todays_open_exceptions = notifications_data[
    pd.to_datetime(notifications_data['NOTFCN_CRTE_TMS']).dt.date == hardcoded_date_dt
]['NOTFCN_ID'].value_counts().to_dict()

# Generate the report data
report_data = []
for analyst, notifications in assignments.items():
    for notification in notifications:
        open_exception_count = open_exceptions_current_month.get(notification, 0)
        todays_exception_count = todays_open_exceptions.get(notification, 0)
        total_exceptions = open_exception_count + todays_exception_count
        
        report_data.append({
            'Notification': notification,
            'Analyst': analyst,
            'Open Exceptions (for current month)': open_exception_count,
            'Today Open Exception': todays_exception_count,
            'Total Exceptions': total_exceptions
        })

report_df = pd.DataFrame(report_data)

# Save the report
report_file = os.path.join(output_path, f"notification_report_{datetime.today().year}.xlsx")
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
