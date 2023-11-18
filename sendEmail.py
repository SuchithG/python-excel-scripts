import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.application import MIMEApplication
import os

# Email settings
smtp_server = "smtp.office365.com"
smtp_port = 587
username = "your-email@outlook.com"
password = "your-password"

# Create message
message = MIMEMultipart()
message["Subject"] = "Test Email with Excel Attachment"
message["From"] = username
message["To"] = "recipient@example.com"

# Email body
body = "Hi,\n\nPlease find the attached Excel file."
message.attach(MIMEText(body, "plain"))

# File settings
filename = "example.xlsx"  # Replace with your file's name
filepath = "/path/to/your/file/" + filename  # Replace with your file's path

# Attach file
with open(filepath, "rb") as attachment:
    part = MIMEApplication(attachment.read(), Name=os.path.basename(filepath))
    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(filepath)}"'
    message.attach(part)

# Send email
try:
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()  # Secure the connection
        server.login(username, password)
        server.sendmail(username, "recipient@example.com", message.as_string())
        print("Email sent successfully with attachment!")
except Exception as e:
    print(f"Error: {e}")