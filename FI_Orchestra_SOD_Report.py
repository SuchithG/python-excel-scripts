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

    analyst_weeks_path = os.path.join(base_path, 'analyst_weeks.xlsx')
    weekly_assignments_path = os.path.join(base_path, 'weekly_assignments.xlsx')
    attendance_tracker_path = os.path.join(base_path, 'Attendance Tracker.xlsx')
    notifications_data_path = os.path.join(base_path, '28 June.xlsx')

    return analyst_weeks_path, weekly_assignments_path, attendance_tracker_path, notifications_data_path, today, previous_day

# Function to get current week number
def get_week_number():
    return datetime.today().isocalendar()[1]

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

    # Equally distribute unassigned notifications among available analysts
    total_notifications = len(unassigned_notifications)
    if total_notifications > 0:
        notifications_per_analyst = total_notifications // len(available_analysts)
        extra_notifications = total_notifications % len(available_analysts)
        
        for i, analyst in enumerate(available_analysts):
            start_index = i * notifications_per_analyst
            end_index = start_index + notifications_per_analyst
            assignments[analyst].extend(unassigned_notifications[start_index:end_index])
        
        # Distribute the extra notifications
        for i in range(extra_notifications):
            assignments[available_analysts[i]].append(unassigned_notifications[-(i+1)])

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

# Main function to encapsulate the script logic
def main():
    # Load paths
    base_path = 'your/base/path/here'
    analyst_weeks_path, weekly_assignments_path, attendance_tracker_path, notifications_data_path, today, previous_day = get_dynamic_paths(base_path)

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
        new_week_analysts = pd.DataFrame({
            'Analyst': all_analysts,
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

    # Generate the report data
    report_data = []
    for analyst, notifications in assignments.items():
        for notification in notifications:
            report_data.append({
                'Notification': notification,
                'Analyst': analyst
            })

    report_df = pd.DataFrame(report_data)

    # Save the report
    report_file = f"notification_report_{datetime.today().year}.xlsx"
    report_df.to_excel(report_file, index=False)
    print(f"Report has been generated and saved as {report_file}")

    # Send the email with the report
    send_email(
        subject="Daily Exception Report",
        body="Please find attached the daily exception report.",
        to_emails=['suchith.girishkumar@db.com'],
        attachment_path=report_file
    )

# Run the main function
main()
