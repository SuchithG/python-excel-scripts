import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
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
    """Filter data for the previous working day and 'Equity' Asset Class."""
    prev_work_day = previous_working_day()
    return df[(df['Date'] == prev_work_day) & (df['Region'] == 'EQ')]


def send_email_with_table(subject, body, recipients, cc_recipients, file_path):
    # Set up the email server and login
    server = smtplib.SMTP('', )
    server.starttls()
    server.login("your_email@gmail.com", 'password')
    from_email = "your_password"  # Consider using an app-specific password if using Gmail

    # Create the email message
    msg = MIMEMultipart()
    msg["From"] = ""
    msg["To"] = ', '.join(recipients)
    msg["Cc"] = ', '.join(cc_recipients)
    msg["Subject"] = subject


    # Attach the body and the excel file to email
    msg.attach(MIMEText(body, 'plain'))
    with open(file_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {file_path.split('/')[-1]}")
        msg.attach(part)

    # Send the email
    all_recipients = recipients + cc_recipients
    server.sendmail(from_email, all_recipients, msg.as_string( ))
    server.quit()
    return f"Email sent to {', '.join(all_recipients)}!"

def add_total_row(df, columns_to_sum):
    # Calculate the total for each column and create a total row
    total_row = {column: df[column].sum() if column in columns_to_sum else 'Total' for column in df.columns}
    total_row[df.columns[0]] = 'Total'  # Set 'Total' label in the first column

    # Adjust for non-summable (non-numeric) columns, set them to 'Total' or some other appropriate string
    for column in df.columns:
        if column not in columns_to_sum:
            total_row[column] = 'Total'

    # Create a DataFrame of the total row
    total_row_df = pd.DataFrame([total_row])

    # Ensure that the total for each numeric column is an integer
    for column in columns_to_sum:
        total_row_df[column] = total_row_df[column].astype(int)

    # Concatenate the total row DataFrame to the original DataFrame
    df_with_total = pd.concat([df, total_row_df], ignore_index=True)

    return df_with_total 

def process_and_send_email_with_tables():
    # Load and filter data
    df = pd.read_excel("/path/to/your/excel_file.xlsx")

    # Convert 'Date' column to datetime format
    df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y').dt.date

    # Filter data for previous working day and 'APAC' region
    filtered_data = filtered_data_for_previous_working_day(df)

    # If no data for the previous working day, print a statement and exit
    if filtered_data.empty:
        return f"No data available for the previous working day ({prev_work_day})."
    
    # Filter for 'Process Activity' in the 'Category' column
    filtered_data_process_activity = filtered_data[filtered_data['Category'] == 'Process Activity']

    # Calculate the '2 eye count' table for the filtered data
    aggregated_by_process_activity = filtered_data_process_activity.groupby(['Date', 'Work Drivers', 'Category', 'Activity', 'Asset Class']).agg({
        'Count': 'sum',
        'Setup': 'sum',
        'Amend': 'sum',
        'Review': 'sum',
        'Closure': 'sum',
        'Deletion': 'sum',
        '4 eye Count': 'sum',
        'Error Count': 'sum',
    }).reset_index()
    aggregated_by_process_activity['Total Count'] = aggregated_by_process_activity[['Count','Setup','Amend','Review','Closure','Deletion','4 eye Count','Error Count']].sum(axis=1)
    columns_to_sum = ['Count','Setup','Amend','Review','Closure','Deletion','4 eye Count','Error Count']
    aggregated_by_process_activity_with_total = add_total_row(aggregated_by_process_activity, columns_to_sum)

    # Filter for 'Proactive Checks' in the 'Category' column
    filtered_data_proactive_checks = filtered_data[filtered_data['Category'] == 'Proactive Checks']

    # Calculate the '2 eye count' table for the filtered data
    aggregated_by_proactive_checks = filtered_data_proactive_checks.groupby(['Date', 'Work Drivers', 'Category', 'Activity', 'Asset Class']).agg({
        'Count': 'sum',
        'Setup': 'sum',
        'Amend': 'sum',
        'Review': 'sum',
        'Closure': 'sum',
        'Deletion': 'sum',
        '4 eye Count': 'sum',
        'Error Count': 'sum',
    }).reset_index()
    aggregated_by_proactive_checks['Total Count'] = aggregated_by_proactive_checks[['Count','Setup','Amend','Review','Closure','Deletion','4 eye Count','Error Count']].sum(axis=1)
    columns_to_sum = ['Count','Setup','Amend','Review','Closure','Deletion','4 eye Count','Error Count']
    aggregated_by_proactive_checks_with_total = add_total_row(aggregated_by_proactive_checks, columns_to_sum)

    # Calculate the 'Count by Resource Name' table
    aggregated_by_resource_name = filtered_data.groupby(['Date', 'Resource Name', 'Asset Class']).agg({
        'Count': 'sum',
        'Setup': 'sum',
        'Amend': 'sum',
        'Review': 'sum',
        'Closure': 'sum',
        '4 eye Count': 'sum',
        'Error Count': 'sum',
    }).reset_index()
    aggregated_by_resource_name['Total Count'] = aggregated_by_resource_name[['Count','Setup','Amend','Review','Closure','Deletion','4 eye Count','Error Count']].sum(axis=1)
    columns_to_sum = ['Count','Setup','Amend','Review','Closure','Deletion','4 eye Count','Error Count']
    aggregated_by_resource_name_with_total = add_total_row(aggregated_by_resource_name, columns_to_sum)

    # Convert columns to integers for all tables 
    all_dfs = [aggregated_by_process_activity, aggregated_by_proactive_checks, aggregated_by_resource_name]
    for df in all_dfs:
        for col in df.columns[2:]:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].fillna(0).astype(int)
    
    def df_to_html_with_integers(df):
        return df.to_html(index=False, float_format=lambda x: '%10.0f' % x)

    # Generate HTML tables for both
    table_by_process_activity_html = df_to_html_with_integers(aggregated_by_process_activity_with_total) if not aggregated_by_process_activity.empty else "<p>No data available for previous working day</p>"
    table_by_proactive_checks_html = df_to_html_with_integers(aggregated_by_proactive_checks_with_total) if not aggregated_by_proactive_checks.empty else "<p>No data available for previous working day</p>"
    table_by_resource_name_html = df_to_html_with_integers(aggregated_by_resource_name_with_total) if not aggregated_by_resource_name.empty else "<p>No data available for previous working day</p>"

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
            <p>Hi Team,</p>
            <p><b>EQ Orchestra Count By <span style="color: Blue;">Process Activity:</span> TCS+DBOI </b> for {prev_work_day}:</p>
            {table_by_process_activity_html}
            <p><b>EQ Orchestra Count By <span style="color: Blue;">Proactive Checks Activity:</span> TCS+DBOI </b> for {prev_work_day}:</p>
            {table_by_proactive_checks_html}
            <p><b>EQ Orchestra Count By <span style="color: Blue;">Resource Name:</span> TCS+DBOI </b> for {prev_work_day}:</p>
            {table_by_resource_name_html}
            <p>Thanks and Regards,</p>
            <p>Suchith Girishkumar</p>
        </body>
    </html>
    """

    # Send the email with all tables
    recepients = ["recipient1_email@example.com","recipient2_email@example.com"]
    cc_recipients = ["recipient1_email@example.com","recipient2_email@example.com"]
    subject = "EQ Consolidated Volumes | Script Testing"
    return send_email_with_table(subject, body, recepients, cc_recipients, r"/path/to/excel_workbook.xlsx")

# Execute the process and print the result
result = process_and_send_email_with_tables()
print(result)
