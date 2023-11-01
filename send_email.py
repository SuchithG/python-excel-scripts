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

def filtered_data_for_previous_working_day(df):
    """Filter data for the previous working day and 'APAC' region."""
    prev_work_day = previous_working_day()
    return df[(df['Date'] == prev_work_day) & (df['Region'] == 'APAC')]


def send_email_with_table(subject, df, body, to_email, attachment_path):
    from_email = "your_email@gmail.com"
    password = "your_password"  # Consider using an app-specific password if using Gmail

    # Convert DataFrame to HTML table
    table_html = df.to_html()
    msg = MIMEMultipart()
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Subject"] = subject

    body = f"""
    <html>
        <head></head>
        <body>
            <p>Hi,</p>
            <p>Here's the "2 eye count" table for {previous_working_day}:</p>
            {table_html}
            <p>Regards,</p>
            <p>Your Name</p>
        </body>
    </html>
    """

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

    # Load and filter data
    df = pd.read_excel("/path/to/your/excel_file.xlsx")
    filtered_data = filtered_data_for_previous_working_day(df)

    # Calculate the '2 eye count' table for the filtered data
    aggregated_data = filtered_data.groupby(['Date', '2 eye']).agg({
        'Setup': 'sum',
        'Amend': 'sum',
        'Closure': 'sum',
        'Deletion': 'sum',
        'Exceptions': 'sum',
    }).reset_index()
    aggregated_data.rename(columns={'2 eye': 'Name'}, inplace=True)
    aggregated_data['2 eye Count'] = aggregated_data[['Setup', 'Amend', 'Closure', 'Deletion', 'Exceptions']].sum(axis=1)

    # Send the email
    response = send_email_with_table("Subject of Email", aggregated_data, "recipient_email@example.com", "/path/to/excel_workbook.xlsx")
    print(response)
