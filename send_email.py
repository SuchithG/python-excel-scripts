import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from datetime import datetime, timedelta

def previous_working_day(today=None):
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

def send_email_with_table(subject, body, recipients, cc_recipients, file_path):
    # Set up the email server and login 
    server = smtplib.SMTP('localhost',34)
    server.starttls()
    server.login('your_email@gmail.com', 'your_password')
    from_email = 'your_email@gmail.com'

    # Create the email message
    msg = MIMEMultipart()
    msg["From"] = "your_email@gmail.com"
    msg["To"] = ', '.join(recipients)
    msg["Cc"] = ', '.join(cc_recipients)
    msg["Subject"] = subject

    # Attach the body and the excel file to email
    msg.attach(MIMEText(body, 'html'))
    with open(file_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {file_path.split('/')[-1]}")
        msg.attach(part)

    # Send the email
    all_recipients = recipients + cc_recipients
    server.sendmail(from_email, all_recipients, msg.as_string())
    server.quit()
    return f"Email sent to {', '.join(all_recipients)}!"

def add_total_row(df, columns_to_sum):
    # Calculate the total for each column and create a total row
    total_row = {column: df[column].sum() if column in columns_to_sum else '' for column in df.columns}
    total_row[df.columns[0]] = 'Total'  # Set 'Total' label in the first column

    # Create a DataFrame of the total row
    total_row_df = pd.DataFrame([total_row])

    # Concatenate the total row DataFrame to the original DataFrame
    df_with_total = pd.concat([df, total_row_df], ignore_index=True)

    # Convert all numeric columns to integers, ignoring non-numeric columns
    for column in columns_to_sum:
        df_with_total[column] = df_with_total[column].astype(int)

    return df_with_total

def process_and_send_email():
        # Load and filter data
        df = pd.read_excel("/path/to/your/excel_file.xlsx")

        # Convert "Date" column to datetime format
        df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y').dt.date

        # Filter date for previous working day and 'APAC' region 
        filtered_data = filtered_data_for_previous_working_day(df)

        # If no data for the 
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
        columns_to_sum = ['Setup', 'Amend', 'Closure', 'Deletion', 'Exceptions', '2 eye Count']
        aggregated_2_eye_data_with_total = add_total_row(aggregated_2_eye_data, columns_to_sum)

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
        columns_to_sum = ['Setup', 'Amend', 'Closure', 'Deletion', 'Exceptions', '2 eye Count']
        aggregated_4_eye_data_with_total = add_total_row(aggregated_4_eye_data, columns_to_sum)

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
        columns_to_sum = ['Setup', 'Amend', 'Closure', 'Deletion', 'Exceptions', '2 eye Count']
        aggregated_by_application_with_total = add_total_row(aggregated_by_application, columns_to_sum)

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
        table_2_eye_html = aggregated_2_eye_data_with_total.to_html(index=False)
        table_4_eye_html = aggregated_4_eye_data_with_total.to_html(index=False)
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
