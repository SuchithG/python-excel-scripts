import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from datetime import datetime, timedelta

def send_email_with_table(subject, body, recipients, cc_recipients, file_path):
    # Set up the email server and login
    server = smtplib.SMTP('', )
    server.starttls()
    server.login("your_email@gmail.com")
    from_email = "your_password" 

    # Convert the email message
    msg = MIMEMultipart()
    msg["From"] = "your_email"
    msg["To"] = ', '.join(recipients)
    msg['Cc'] = ', '.join(cc_recipients)
    msg["Subject"] = subject

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
    total_row[df.columns[0]] = 'Total' # Set 'Total' label in the first column 
    # Create a DataFrame of the total row
    total_row_df = pd.DataFrame([total_row])
    # Concatenated the total row DataFrame to the original DataFrame
    df_with_total = pd.concat([df, total_row_df], ignore_index=True)
    # Convert all numeric columns to integers, ignoring non-numeric columns
    for column in columns_to_sum:
        df_with_total[column] = df_with_total[column].astype(int)
    return df_with_total 

def process_and_send_email_with_tables():
    # Determine the previous month and year for file naming and data filtering
    current_date = datetime.now()
    first_day_of_current_month = datetime(current_date.year, current_date.month, 1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    previous_month_name = last_day_of_previous_month.strftime("%b-%y")  # Format: Nov-23

    # Construct file names
    input_file_name = f"ODC_ConsolidatedFile_Monthly_{previous_month_name}.xlsx"
    output_file_name = f"ODC_ConsolidatedFile_Monthly_{previous_month_name}.xlsx"

    # Load data from the primary Excel file
    df = pd.read_excel(input_file_name)

    # Load and filter data from the secondary "Summary Sheet"
    summary_df = pd.read_excel("path_to_second_excel_file.xlsx", sheet_name="Summary Sheet")
    summary_df['Date'] = pd.to_datetime(summary_df['Date'])

    # Filter for the specific month and year
    summary_df_filtered = summary_df[(summary_df['Date'].dt.month == 11) &  # Example month
                                     (summary_df['Date'].dt.year == 2023)]  # Example year
    if summary_df_filtered.empty:
        print("No data for November 2023.")
    else:
        print("Data available for November 2023.")

    # Calculate the BOT Volumes
    columns_to_sum_bot = ['Asset Class/Reports', 'Application', 'Setup', 'Amend', 'Review', 'Closure', 'Deletion', 'Exceptions']
    bot_volumes = summary_df_filtered[columns_to_sum_bot].sum().sum()

    # Remove duplicate entries in 'PDF Name'
    unique_pdf_df = df.drop_duplicates(subset=['PDF Name'])

    # Convert 'Date' column to datetime format
    df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y').dt.date

    aggregated_data_1 = df.groupby(['Region']).agg({
        'Setup': 'sum',
        'Amend': 'sum',
        'Closure': 'sum',
        'Deletion': 'sum',
        'Exceptions': 'sum',
        'PDF Name': 'count'
    }).reset_index()
    aggregated_data_1.rename(columns={'PDF Name': 'PDF Count'}, inplace=True)
    columns_to_sum = ['Setup', 'Amend', 'Closure', 'Deletion', 'Exceptions', 'PDF Count']
    aggregated_data_with_total_1 = add_total_row(aggregated_data_1, columns_to_sum)

    aggregated_data_2 = df.groupby(['Region']).agg({
        'PDF missed(late 4 eye/stamping)': 'sum',
        'Error Count': 'sum',
    }).reset_index()
    aggregated_data_2.rename(columns={'PDF missed(late 4 eye/stamping)': 'PDF SLA Missed Count', 'Error Count':'4-eye Error Count'}, inplace=True)
    columns_to_sum = ['PDF SLA Missed Count', '4-eye Error Count']
    aggregated_data_with_total_2 = add_total_row(aggregated_data_2, columns_to_sum)

     # Aggregation for the new table
    total_pdf_count = len(unique_pdf_df['PDF Name'])
    total_security_setup = df['Setup'].sum()
    total_security_amendments = df[['Amend', 'Closure', 'Deletion']].sum().sum()  # Sum across columns, then total sum
    total_security_review = df['Review'].sum()
    total_espear_exceptions = df['Exceptions'].sum()

    # Creating a DataFrame for the new table
    data_for_new_table = {
        'PDF Count': [total_pdf_count],
        'Security Setup': [total_security_setup],
        'Security Amendments': [total_security_amendments],
        'Security Review': [total_security_review],
        'Espear Exceptions': [total_espear_exceptions],
        'Bot Volumes': [bot_volumes]
    }
    new_table_df = pd.DataFrame(data_for_new_table)

    # Convert numeric values to integers in new_table_df
    for column in data_for_new_table:
        new_table_df[column] = pd.to_numeric(new_table_df[column], errors='coerce').fillna(0).astype(int)

    # Generate HTML tables for both
    table_html_1 = aggregated_data_with_total_1.to_html(index=False) if not aggregated_data_1.empty else "<p>No data available</p>"
    table_html_2 = aggregated_data_with_total_2.to_html(index=False) if not aggregated_data_2.empty else "<p>No data available</p>"
    table_html_new = new_table_df.to_html(index=False)

    body = f"""
    <html>
        <head>
            <style>
            table, th, td {{
                border: 1px solid black;
                border-collapse: collapse;
                text-align: center; /* Center-align text*/
                padding: 5px; /* Optional: to add some padding inside cells */
            }}
        </style>
    </head>
        <body>
            <p >Hi,</p>
            <p>Enclosed is the ODC Monthly Volumes for the month - {previous_month_name}.</p>
            <p><b>Total count</b><p>
            {table_html_1}
            <p><b>Total count</b></p>
            {table_html_2}
            <p><b>Total count</b></p>
            {table_html_new}
            <p>Thanks and Regards,</p>
            <p>Suchith</p>
        </body>
    </html>
    """

    # Send the email with all tables
    recipients = [""]
    cc_recipients = [""]
    subject = f"ODC Monthly Volumes - {previous_month_name} | Script Testing"
    return send_email_with_table(subject, body, recipients, cc_recipients, output_file_name)

# Execute the process and print the result
result = process_and_send_email_with_tables()
print(result)