import os
import cx_Oracle
import pandas as pd
from datetime import datetime, timedelta

# Function to get the previous month's dates in the desired format
def get_previous_month_dates():
    today = datetime.today()
    first_day_current_month = today.replace(day=1)
    last_day_prev_month = first_day_current_month - timedelta(days=1)
    first_day_prev_month = last_day_prev_month.replace(day=1)
    
    previous_month_start = first_day_prev_month.strftime("%d-%b-%Y").upper()
    previous_month_end = last_day_prev_month.strftime("%d-%b-%Y").upper()

    return previous_month_start, previous_month_end

# Function to execute queries and save data to excel
def execute_queries_and_save_to_excel(connection, queries, file_path, previous_month_start, previous_month_end):
    print("Fetching data and storing in Excel...")
    # Create a Pandas Excel writer
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as excel_writer:
        for query_name, query in queries.items():
            # Replace placeholders in the query
            query = query.replace(":prev_month_start", f"'{previous_month_start}'")
            query = query.replace(":prev_month_end", f"'{previous_month_end}'")
            
            # Create a valid sheet name (Excel has a 31 char limit for sheet names)
            sheet_name = query_name[:31]
            
            try:
                # Fetch the query results into a DataFrame
                df = pd.read_sql(query, con=connection)
                
                # Write the DataFrame to the Excel file using the sheet name
                df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
                print(f"Data stored for query: {query_name}")
            except cx_Oracle.DatabaseError as e:
                error = e.args
                print(f"Oracle Database error for query {query_name}:", error)
        
        # Save the Excel file (This is done automatically when exiting the 'with' block)
        print("Data saved to Excel successfully.")

# Connection details
username = ""
password = ""
dsn = ""

# Excel folder path
excel_folder_path = r""

# Ensure the directory exists or create it if it doesn't
if not os.path.exists(excel_folder_path):
    os.makedirs(excel_folder_path)

# Get the dates for the previous month
previous_month_start, previous_month_end = get_previous_month_dates()

# Excel file name with desired format
file_name = f"DQ_CDE_SOI_{previous_month_start.split('-')[1]}.xlsx"
file_path = excel_folder_path + '\\' + file_name

# Define the queries
queries = {
    # ... [Your queries go here]
}

print("Initializing connection")
try:
    connection = cx_Oracle.connect(user=username, password=password, dsn=dsn)
    # Execute queries and save data to Excel
    execute_queries_and_save_to_excel(connection, queries, file_path, previous_month_start, previous_month_end)
    connection.close()
    print("Connection closed.")
except cx_Oracle.DatabaseError as e:
    error = e.args
    print("Oracle Database error during connection:", error)
