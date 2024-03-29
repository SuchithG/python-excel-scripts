import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

# Function to get the previous month's file name
def get_previous_month_filename():
    today = datetime.today()
    first = today.replace(day=1)
    last_month = first - timedelta(days=1)
    return f"ODC_ConsolidatedFile_Monthly_{last_month.strftime('%b-%y')}.xlsx"

# Dynamically get the file path
excel_file_path = 'path/to/your/files/' + get_previous_month_filename()

# Read the Excel file
df = pd.read_excel(excel_file_path)

# Function to convert columns to integers
def convert_columns_to_int(df, columns):
    for col in columns:
        df[col] = df[col].fillna(0).astype(int)
    return df

# Create Table 1 with 'PDF Name' aggregated as 'PDF Count'
table1_columns = ['Region', 'Setup', 'Amend', 'Review', 'Closure', 'Exceptions', 'PDF Name']
table1 = df[table1_columns].groupby('Region').agg({
    'Setup': 'sum', 
    'Amend': 'sum', 
    'Review': 'sum', 
    'Closure': 'sum', 
    'Exceptions': 'sum', 
    'PDF Name': 'count'  # Counting the occurrences of 'PDF Name'
}).reset_index()
table1.rename(columns={'PDF Name': 'PDF Count'}, inplace=True)  # Rename the column

# Calculate total row
table1_total = pd.DataFrame([table1[['Setup', 'Amend', 'Review', 'Closure', 'Exceptions', 'PDF Count']].sum()], columns=table1.columns[1:])
table1_total['Region'] = 'Total'

# Append the total row
table1 = pd.concat([table1, table1_total], ignore_index=True)

# Convert numeric columns to int
numeric_columns_table1 = ['Setup', 'Amend', 'Review', 'Closure', 'Exceptions', 'PDF Count']
table1 = convert_columns_to_int(table1, numeric_columns_table1)


# Create Table 2
table2_columns = ['Region', 'PDF missed(late 4 eye/stamping)', 'Error Count']
table2 = df[table2_columns].groupby('Region').sum().reset_index()
table2.rename(columns={'PDF missed(late 4 eye/stamping)': 'PDF SLA Missed Count', 'Error Count': '4-eye Error Count'}, inplace=True)  # Rename the columns
table2_total = pd.DataFrame([table2.sum(numeric_only=True)], columns=table2.columns)
table2_total['Region'] = 'Total'
table2 = pd.concat([table2, table2_total], ignore_index=True)

# Convert numeric columns to int
numeric_columns_table2 = ['PDF SLA Missed Count', '4-eye Error Count']
table2 = convert_columns_to_int(table2, numeric_columns_table2)

# Convert tables to HTML
html_table1 = table1.to_html(index=False)
html_table2 = table2.to_html(index=False)

# Email settings
smtp_server = "smtp.office365.com"
smtp_port = 587
username = "your-email@outlook.com"
password = "your-password"

# Define lists of recipients
to_recipients = ["recipient1@example.com", "recipient2@example.com"]
cc_recipients = ["cc1@example.com", "cc2@example.com"]
bcc_recipients = ["bcc1@example.com", "bcc2@example.com"]

# Combine all recipients for the sendmail function
all_recipients = to_recipients + cc_recipients + bcc_recipients

# Create message
message = MIMEMultipart("mixed")
message["Subject"] = "Email with Tables from Python"
message["From"] = username
message["To"] = ", ".join(to_recipients)
message["CC"] = ", ".join(cc_recipients)

# Email body with tables
html_part = MIMEMultipart("alternative")
html = f"""
<html>
  <head>
    <style>
        body {{
            font-family: Arial, sans-serif;
            font-size: 12px; /* Adjust the font size of the body, affecting overall readability */
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            font-size: 10px; /* Smaller font size for the table */
        }}
        th, td {{
            border: 1px solid #dddddd;
            text-align: left;
            padding: 4px; /* Reduced padding inside cells */
        }}
        th {{
            background-color: #f2f2f2;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        tr.total {{
            font-weight: bold;
            background-color: #e2e2e2;
        }}
    </style>
</head>
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
html_part.attach(MIMEText(html, "html"))

# Attach the HTML part
message.attach(html_part)

# Attachment part (for the Excel file)
part = MIMEBase('application', "octet-stream")
with open(excel_file_path, 'rb') as file:
    part.set_payload(file.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment', filename=get_previous_month_filename())
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
