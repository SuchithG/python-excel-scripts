import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Read the Excel file
excel_file = "path/to/your/excel/file.xlsx"
df = pd.read_excel(excel_file)

# Create Table 1
table1_columns = ["Region", "Setup", "Amend", "Review", "Closure", "Exceptions", "PDF Name"]
table1 = df[table1_columns].copy()
table1_total = table1.groupby("Region").sum().reset_index()
table1_html = table1_total.to_html(index=False)

# Create Table 2
table2_columns = ["Region", "PDF missed(late 4 eye/stamping)", "Error Count"]
table2 = df[table2_columns].copy()
table2_total = table2.groupby("Region").sum().reset_index()
table2_html = table2_total.to_html(index=False)

# Email settings
smtp_server = "smtp.office365.com"
smtp_port = 587
username = "your-email@outlook.com"
password = "your-password"

# Create message
message = MIMEMultipart("alternative")
message["Subject"] = "Tables from Excel in Email Body"
message["From"] = username
message["To"] = "recipient@example.com"

# Email body with HTML tables
html = f"""
<html>
  <body>
    <p>Hi,<br>
       Please find below the tables extracted from the Excel file:</p>
    <h2>Table 1</h2>
    {table1_html}
    <h2>Table 2</h2>
    {table2_html}
  </body>
</html>
"""

# Attach HTML to email
part = MIMEText(html, "html")
message.attach(part)

# Send email
try:
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()  # Secure the connection
        server.login(username, password)
        server.sendmail(username, "recipient@example.com", message.as_string())
        print("Email with tables sent successfully!")
except Exception as e:
    print(f"Error: {e}")
