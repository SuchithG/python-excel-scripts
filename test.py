import cx_Oracle
import pandas as pd
from datetime import datetime, timedelta
from threading import Thread
import time

# Function to get the previous month's dates and formatted strings
def get_previous_month_dates():
    today = datetime.today()
    first_day_current_month = today.replace(day=1)
    last_day_prev_month = first_day_current_month - timedelta(days=1)
    first_day_prev_month = last_day_prev_month.replace(day=1)
    
    previous_month_start = first_day_prev_month.strftime("%d-%b-%y").upper()
    previous_month_end = last_day_prev_month.strftime("%d-%b-%y").upper()
    previous_month_name = first_day_prev_month.strftime("%b").upper()

    current_month_start = first_day_current_month.strftime("%d-%b-%y").upper()

    return { 
         'prev_start': previous_month_start,
         'prev_end': previous_month_end,
         'prev_name': previous_month_name,
         'curr_start': current_month_start
     }

# Function to execute queries and save data to Excel
def execute_queries_and_save_to_excel(connection, queries, file_path, prev_start, prev_end, curr_start):
    print("Fetching data and storing in Excel...")
    # Create a Pandas Excel writer
    excel_writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

    for query_name, query in queries.items():
        query = query.replace(":prev_month_start", f"'{prev_start}'")
        query = query.replace(":curr_month_start", f"'{curr_start}'")
        query = query.replace(":prev_month_end", f"'{prev_end}'")
        
        sheet_name = query_name[:31] # Limit sheet name length to 31 characters
        try:
             if query_name == 'OTC Exception' or query_name == 'OTC Universe':
                 # Establish a separate connection with different credentials
                 separate_connection = cx_Oracle.connect(user=Otc_username, password=Otc_pwd, dsn=Otc_dbname)
                 df = pd.read_sql(query, con=separate_connection)
                 separate_connection.close()
             else:
                 df = pd.read_sql(query, con=connection)
                 df.to_excel(excel_writer, sheet_name, index=False)
                 print(f"Data stored for query: {query_name}")
                 time.sleep(1) # Pause for 1 second to simulate the process 
        except cx_Oracle.DatabaseError as e:
                error = e.args
                print("Oracle Database error:", error)
    excel_writer.close()
    print("Data saved to Excel successfully.")

# Function to establish the database connection with timeout
def db_connection_with_timeout(username, password, dsn, timeout):
    connection = None
    result = {}

    def set_connection():
        nonlocal connection
        connection = cx_Oracle.connect(user=username, password=password, dsn=dsn)

    connect_thread = Thread(target=set_connection)
    connect_thread.start()
    connect_thread.join(timeout)

    if connect_thread.is_alive():
        print("Connection timeout occurred. Please check your network configuration.")
        connect_thread.join() # wait for the thread to terminate

    result['connection'] = connection
    return result
    
# Connection details
Otc_username = ""
Otc_pwd = ""
Otc_dbname = ""
username = ""
password = ""
dsn = ""

# Excel folder path
excel_folder_path = r""

# Get the dates and formatted strings for the previous month
dates = get_previous_month_dates()

# Excel file name with desired format
file_name = f"DQ_CDE_SOI_{dates['prev_name']}.xlsx"
file_path = f"{excel_folder_path}\\{file_name}"

# Define the queries
queries = {
    'Trading Hours Universe': "SELECT ...",
    'OTC Exception': "SELECT ...",  
    'OTC Universe': "",
    'Exchange Universe': "",
}

print("Initializing connection")
result = db_connection_with_timeout(username, password, dsn, timeout=60)
connection = result['connection']

if connection:
     #Execute queries and save data to Excel
     execute_queries_and_save_to_excel(connection, queries, file_path, dates['prev_start'], dates['prev_end'], dates['curr_start'])
     connection.close()
     print("Connection closed.")
