import cx_Oracle
import pandas as pd
from datetime import datetime, timedelta
from threading import Thread
import time

# Function to get the previous mont's dates and formatted strings
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
         previous_month_start,
         previous_month_end,
         previous_month_name,
         current_month_start
    }

# Function to execute queries and save data to excel
def execute_queries_and_save_to_excel(connection, queries, file_path, previous_month_start, previous_month_end, current_month_start):
    print("Fetching data and storing in Excel...")
    # Create a Pandas Excel writer
    excel_writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

    for query_name, query in queries.items():
        query = query.replace(":prev_month_start", f"'{previous_month_start}'")
        query = query.replace(":curr_month_start", f"'{current_month_start}'")
        query = query.replace(":prev_month_end", f"'{previous_month_end}'")
        
        sheet_name = query_name[:31] # Limit sheet name length to 31 charaters
        try:
             if query_name == 'OTC Exception' or query_name == 'OTC Universe':
                 # Establish a seperate connection with different credentials
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

    if connect_thread.is_Alive():
        print("Connection timeout occured. Please check your network configuration.")
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
previous_month_start, previous_month_end, current_month_start, previous_month_name = get_previous_month_dates()

# Excel file name with desired format
file_name = f"DQ_CDE_SOI_{previous_month_name}.xlsx"
file_path = excel_folder_path + '\\' + file_name

        # Define the queries
queries = {
            'Trading Hours Universe': "SELECT LAST_CHG_USR_ID, notfcn_id, COUNT(*) FROM gc_owner.ft_T_NTEL WHERE MSG_TYP IN ('ASX_ Security Masters', 'BBEquityCollateral', 'BBGlobalEquity', 'BBGlobalEquityPricing', 'BBGlobalMutualFund', 'BBGloballutualFundPricing', 'BBGlobalWarrants', 'BBGlobalWarrantsPricing', 'BBShortSellingRestriction', 'CGS', 'RTDSEPlus_Equity_REF_CSTM', 'RTDSEPlus_EquityWarrants_REF_CSTM', 'RTDSEPLUS_MutualFund8&MM_REF_CSTM', 'SMChanges', 'SwapsTradingHours', 'TR_DS_FATCA_Bulk_Solution', 'TR_DS_FATCA_Bulk_Solution DELTA', 'DBTKFSD', 'DBTKFAD', 'Lipper_Holding', 'Lipper_Fund', 'DWS_Holding', 'DWS_Fund', 'FINFRAG_TRADE') AND SOURCE_ID IN ('GC_OWNERGLNORS1P_APP.UK.DB.COM', 'TRANSLATION', 'PROSReconciliation') AND LST_NOTFCN_TMS >= TO_DATE(:prev_month_start, 'DD-MON-YYYY') AND LST_NOTFCN_TMS < TO_DATE(:curr_month_start, 'DD-MON-YYYY') AND NOTFCN_STAT_TYP = 'CLOSED' AND NOTFCN_ID NOT IN ('2', '5', '23', '527', '541', '640', '655', '665', '667', '50001', '50009', '50016', '50021', '50050', '50052', '60202', '60981', '70006', '70007', '70085', '70086', '70087', '16', '67501', '153') GROUP BY LAST_CHG_USR_ID, notfcn_id",
            'OTC Exception': "SELECT TRUNC(NOTFCN_CRTE_TMS), TRUNC(LST_NOTFCN_TMS), notfcn_id, NOTFCN_STAT_TYP, COUNT(*) FROM gc_owner.ft_T_NTEL WHERE MSG_TYP IN ('ASX_Security_Masters', 'BBEquityCollateral', 'BBGlobalEquity', 'BBGlobalEquity Pricing', 'BBGlobalMutualFund', 'BBGlobalMutual Fund Pricing', 'BBGlobalWarrants', 'BBGlobalWarrantsPricing', 'BBShortSellingRestriction', 'CGS', 'RTDSEPlus_Equity_REF_CSTM', 'RTDSEPlus_EquityWarrants_REF_CSTM', 'RTDSEPlus_Mutual Fund&&MM_REF_CSTM', 'SwapsTradingHours', 'TR_DS_FATCA_Bulk_Solution', 'TR_DS_FATCA_Bulk_Solution_DELTA', 'DBTKFSD', 'DBTKFAD', 'Lipper_Holding', 'Lipper_Fund','D WS_Holding', 'DWS_Fund','FINFRAG_TRADE') AND SOURCE_ID IN ('GC_OWNER@LNORS1P_APP.UK.DB.COM','TRANSLATION','PRDSReconciliation') AND TRUNC(NOTFCN_CRTE_TMS) >= TO_DATE(:prev_month_start, 'DD-MON-YY') AND TRUNC(NOTFCN_CRTE_TMS) < TO_DATE(:curr_month_start, 'DD-MON-YY') AND NOTFCN_STAT_TYP IN ('OPEN','ASSIGN') AND NOTFCN_ID NOT IN ('2','5','23','527','541','640','655','665','667','50001','50009','50016','50021','50050','50052','60202','60981','70006','70007','70085','70086','70087','16','67501','153') GROUP BY TRUNC(NOTFCN_CRTE_TMS), TRUNC(LST_NOTFCN_TMS), notfcn_id, NOTFCN_STAT_TYP",  
            'OTC Universe': "" ,
            'Exchange Universe': "" ,
        }


print("Initializing connection")
result = db_connection_with_timeout(username, password, dsn, timeout=60)
connection = result['connection']

if connection:
     #Execute queries and save data to Excel
     execute_queries_and_save_to_excel(connection, queries, file_path, previous_month_start, previous_month_end, current_month_start)
     connection.close()
     print("Connection closed.")