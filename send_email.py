import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email import encoders
import pandas as pd

# Load your data (assuming you have it in an Excel file)
df = pd.read_excel("/path/to/your/excel_file.xlsx")

# Process your data to get the '2 eye count' table (as we did earlier)
aggregated_data = df.groupby(['Date', '2 eye']).agg({
    'Setup': 'sum',
    'Amend': 'sum',
    'Closure': 'sum',
    'Deletion': 'sum',
    'Exceptions': 'sum',
}).reset_index()
aggregated_data.rename(columns={'2 eye': 'Name'}, inplace=True)
aggregated_data['2 eye Count'] = aggregated_data[['Setup', 'Amend', 'Closure', 'Deletion', 'Exceptions']].sum(axis=1)

def send_email_with_table(subject, body, to_email, attachment_path):
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
            <p>Here's the "2 eye count" table:</p>
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

# Send the email
response = send_email_with_table("Subject of Email", aggregated_data, "recipient_email@example.com", "/path/to/excel_workbook.xlsx")
print(response)
