import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from datetime import datetime, timedelta

def previous_working_day(today=None, df=None):
    """Compute the previous working day"""
    if today is None:
        today = datetime.now().date()

        offset = 1 if today.weekday() != 0 else 3
        return today - timedelta(days=offset)

# Define prev_work_day as a global variable
prev_work_day = previous_working_day()

def filtered_data_for_previous_working_day(df):
    """Filter data for the previous working day and 'APAC' region."""
    prev_work_day = previous_working_day()
    return df[(df['Date'] == prev_work_day) & (df['Region' == 'APAC'])]

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
    try:
        df = pd.read_excel("/path/to/your/excel_file.xlsx")
        df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y').dt.date
    except Exception as e:
        return f"Error loading data: {e}"

    filtered_data = filtered_data_for_previous_working_day(df)
    if filtered_data is None or filtered_data.empty:
        return f"No data available for the previous working day."

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

    # Calculate the 'Count by application and asset class' table
    aggregated_by_app_and_asset_class = filtered_data.groupby(['Date', 'Application', 'Asset Class/Reports']).agg({
        'Setup': 'sum',
        'Amend': 'sum',
        'Review': 'sum',
        'Closure': 'sum',
        'Deletion': 'sum',
        'Exceptions': 'sum',
    }).reset_index()
    aggregated_by_app_and_asset_class['Total Count'] = aggregated_by_app_and_asset_class[['Setup', 'Amend', 'Review', 'Closure', 'Deletion', 'Exceptions']].sum(axis=1)

    # Convert columns to integers for all tables 
    all_dfs = [aggregated_2_eye_data, aggregated_4_eye_data, aggregated_by_application, aggregated_by_app_and_asset_class]
    for df in all_dfs:
        for col in df.columns[2:]:
            if df[col].dtype == 'float64':  # Ensure only numeric columns undergo the conversion
                df[col] = df[col].fillna(0).astype(int)

    # Generate HTML tables for both
    table_2_eye_html = aggregated_2_eye_data.to_html(index=False)
    table_4_eye_html = aggregated_4_eye_data.to_html(index=False)
    table_by_application_html = aggregated_by_application.to_html(index=False)
    table_by_app_and_asset_class_html = aggregated_by_app_and_asset_class.to_html(index=False)

    body = f"""
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
            <p style="text-align:center;">Hi,</p>
            <p style="text-align:center;">Here's the "2 eye count" table for {previous_working_day}:</p>
            {table_2_eye_html}
            <p style="text-align:center;">Here's the "2 eye count" table for {previous_working_day}:</p>
            {table_4_eye_html}
            <p style="text-align:center;">Here's the "Count by application" table for {prev_work_day}:</p>
            {table_by_application_html}
            <p style="text-align:center;">Here's the "Count by application and asset class" table for {prev_work_day}:</p>
            {table_by_app_and_asset_class_html}
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
