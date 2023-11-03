import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email import encoders
import pandas as pd
from datetime import datetime, timedelta

def previous_working_day(today=None):
    """Compute the previous working day."""
    if today is None:
        today = datetime.now().date()
    
    offset = 1 if today.weekday() != 0 else 3
    return today - timedelta(days=offset)

# Define prev_work_day as a global variable
prev_work_day = previous_working_day()

def filtered_data_for_previous_working_day(df):
    """Filter data for the previous working day and 'APAC' region."""
    prev_work_day = previous_working_day()
    return df[(df['Date'] == prev_work_day) & (df['Region'] == 'APAC')]


def send_email_with_table(subject, df, body, to_email, attachment_path):
    from_email = "your_email@gmail.com"
    password = "your_password"  # Consider using an app-specific password if using Gmail

    # Convert DataFrame to HTML table
    table_html = df.to_html(index=False)
    msg = MIMEMultipart()
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Subject"] = subject



    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {attachment_path}")
        msg.attach(part)

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(from_email, password)
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Error occurred: {e}")

def process_and_send_email():
    # Load and filter data
    df = pd.read_excel("/path/to/your/excel_file.xlsx")

    # Convert 'Date' column to datetime format
    df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y').dt.date

    # Filter data for previous working day and 'APAC' region
    filtered_data = filtered_data_for_previous_working_day(df)

    # If no data for the previous working day, print a statement and exit
    if filtered_data.empty:
        return f"No data available for the previous working day ({prev_work_day})."

    # Calculate the '2 eye count' table for the filtered data
    aggregated_2_eye_data = filtered_data.groupby(['Date', '2 eye']).agg({
        'Setup': 'sum',
        'Amend': 'sum',
        'Closure': 'sum',
        'Deletion': 'sum',
        'Exceptions': 'sum',
    }).reset_index()
    aggregated_2_eye_data.rename(columns={'2 eye': 'Name'}, inplace=True)
    aggregated_2_eye_data['2 eye Count'] = aggregated_2_eye_data[['Setup', 'Amend', 'Closure', 'Deletion', 'Exceptions']].sum(axis=1)

    # Calculate the '4 eye count' table
    aggregated_4_eye_data = filtered_data.groupby(['Date', '4 eye']).agg({
        'Setup': 'sum',
        'Amend': 'sum',
        'Review': 'sum',
        'Closure': 'sum',
        'Deletion': 'sum',
        'Exceptions': 'sum',
    }).reset_index()
    aggregated_4_eye_data.rename(columns={'4 eye': 'Name'}, inplace=True)
    aggregated_4_eye_data['4 eye Count'] = aggregated_4_eye_data[['Setup', 'Amend', 'Review', 'Closure', 'Deletion', 'Exceptions']].sum(axis=1)

    # Calculate the 'Count by application' table
    aggregated_by_application = filtered_data.groupby(['Date', 'Application']).agg({
        'Setup': 'sum',
        'Amend': 'sum',
        'Review': 'sum',
        'Closure': 'sum',
        'Deletion': 'sum',
        'Exceptions': 'sum',
    }).reset_index()
    aggregated_by_application['Total Count'] = aggregated_by_application[['Setup', 'Amend', 'Review', 'Closure', 'Deletion', 'Exceptions']].sum(axis=1)

    # Convert columns to integers for all tables
    for df in [aggregated_2_eye_data, aggregated_4_eye_data, aggregated_by_application]:
        for col in df.columns[2:]:
            df[col] = df[col].astype(int)

    # Generate HTML tables for both
    table_2_eye_html = aggregated_2_eye_data.to_html(index=False)
    table_4_eye_html = aggregated_4_eye_data.to_html(index=False)
    table_by_application_html = aggregated_by_application.to_html(index=False)

    body = f"""
    <html>
        <head>
        <style>
            table {{
                width: 100%;
                border-collapse: collapse;
            }}
            table, th, td {{
                border: 1px solid black;
                text-align: center;
                vertical-align: middle;
                white-space: normal;
            }}
        </style>
    </head>
        <body>
            <p style="text-align:center;">Hi,</p>
            <p style="text-align:center;">Here's the "2 eye count" table for {previous_working_day}:</p>
            {table_2_eye_html}
            <p style="text-align:center;">Here's the "2 eye count" table for {previous_working_day}:</p>
            {table_4_eye_html}
            <p style="text-align:center;">Here's the "Count by application" table for {prev_work_day}:</p>
            {table_by_application_html}
            <p style="text-align:center;">Regards,</p>
            <p style="text-align:center;">Your Name</p>
        </body>
    </html>
    """

    # Send the email
    return send_email_with_table("Subject of Email", body, "recipient_email@example.com", "/path/to/excel_workbook.xlsx")

# Execute the process and print the result
result = process_and_send_email()
print(result)
