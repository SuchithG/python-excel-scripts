import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Read the Excel file
df = pd.read_excel('path/to/your/excel/file.xlsx')

# Create Table 1
table1_columns = ['Region', 'Setup', 'Amend', 'Review', 'Closure', 'Exceptions', 'PDF Name']
table1 = df[table1_columns].groupby('Region').sum().reset_index()
table1_total = pd.DataFrame([table1.sum(numeric_only=True)], columns=table1.columns)
table1_total['Region'] = 'Total'
table1 = pd.concat([table1, table1_total], ignore_index=True)

# Create Table 2
table2_columns = ['Region', 'PDF missed(late 4 eye/stamping)', 'Error Count']
table2 = df[table2_columns].groupby('Region').sum().reset_index()
table2_total = pd.DataFrame([table2.sum(numeric_only=True)], columns=table2.columns)
table2_total['Region'] = 'Total'
table2 = pd.concat([table2, table2_total], ignore_index=True)

# Convert tables to HTML
html_table1 = table1.to_html(index=False)
html_table2 = table2.to_html(index=False)

# Email settings
smtp_server = "smtp.office365.com"
smtp_port = 587
username = "your-email@outlook.com"
password = "your-password"

# Create message
message = MIMEMultipart("alternative")
message["Subject"] = "Email with Tables from Python"
message["From"] = username
message["To"] = "recipient@example.com"

# Email body with tables
html = f"""
<html>
  <head></head>
  <body>
    <p>Hi,<br>
       Please find below the required tables:<br>
       <h3>Table 1:</h3>
       {html_table1}
       <h3>Table 2:</h3>
       {html_table2}
    </p>
  </body>
</html>
"""
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
