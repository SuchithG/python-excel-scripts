import os
import cx_Oracle
import pandas as pd
from datetime import datetime, timedelta
import time

def get_previous_month_dates():
    today = datetime.today()
    first_day_prev_month = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
    first_day_current_month = today.replace(day=1)

    first_day_prev_month_str = first_day_prev_month.strftime("%d-%b-%Y").upper()
    first_day_current_month_str = first_day_current_month.strftime("%d-%b-%Y").upper()

    previous_month = first_day_prev_month
    previous_month_name = previous_month.strftime("%b").upper()
    previous_month_year = previous_month.strftime("%Y")

    return (
        first_day_prev_month_str,
        first_day_current_month_str,
        previous_month_name,
        previous_month_year,
    )

# Function to execute queries and save data to excel
def execute_queries_and_save_to_excel(connection, queries, file_path, first_day_prev_month_str, first_day_current_month_str):
    print("Fetching data and storing in Excel...")
    # Create a Pandas Excel writer
    excel_writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

    for query_name, query in queries.items():
        query = query.replace(":prev_month_start", f"'{first_day_prev_month_str}'")
        query = query.replace(":curr_month_start", f"'{first_day_current_month_str}'")
        df = pd.read_sql(query, con=connection)
        df.to_excel(excel_writer, sheet_name=query_name, index=False)
        print(f"Data stored for query: {query_name}")
        time.sleep(1)

    # Save the Excel file
    excel_writer.close()

    print("Data saved to Excel successfully.")

def db_connection(username, pwd, dbname, excel_folder_path):
    connection = None

    try:
        connection = cx_Oracle.connect(user=username, password=pwd, dsn=dbname)
        print("Connected to Oracle Database:", connection.version)

        (
            first_day_prev_month_str,
            first_day_current_month_str,
            previous_month_name,
            previous_month_year,
        ) = get_previous_month_dates()

        # Excel file name with desired format
        file_name = f"iRDS_SQL_QUERY_{previous_month_name}_ME_{previous_month_year}.xlsx"
        file_path = os.path.join(excel_folder_path, file_name)

        queries = {
            'Line 665': f"SELECT LAST_CHG_USR_ID, notfcn_id, COUNT(*) FROM gc_owner.ft_T_NTEL WHERE MSG_TYP IN ('ASX_ Security Masters', 'BBEquityCollateral', 'BBGlobalEquity', 'BBGlobalEquityPricing', 'BBGlobalMutualFund', 'BBGloballutualFundPricing', 'BBGlobalWarrants', 'BBGlobalWarrantsPricing', 'BBShortSellingRestriction', 'CGS', 'RTDSEPlus_Equity_REF_CSTM', 'RTDSEPlus_EquityWarrants_REF_CSTM', 'RTDSEPLUS_MutualFund8&MM_REF_CSTM', 'SMChanges', 'SwapsTradingHours', 'TR_DS_FATCA_Bulk_Solution', 'TR_DS_FATCA_Bulk_Solution DELTA', 'DBTKFSD', 'DBTKFAD', 'Lipper_Holding', 'Lipper_Fund', 'DWS_Holding', 'DWS_Fund', 'FINFRAG_TRADE') AND SOURCE_ID IN ('GC_OWNERGLNORS1P_APP.UK.DB.COM', 'TRANSLATION', 'PROSReconciliation') AND LST_NOTFCN_TMS >= TO_DATE(:prev_month_start, 'DD-MON-YYYY') AND LST_NOTFCN_TMS < TO_DATE(:curr_month_start, 'DD-MON-YYYY') AND NOTFCN_STAT_TYP = 'CLOSED' AND NOTFCN_ID NOT IN ('2', '5', '23', '527', '541', '640', '655', '665', '667', '50001', '50009', '50016', '50021', '50050', '50052', '60202', '60981', '70006', '70007', '70085', '70086', '70087', '16', '67501', '153') GROUP BY LAST_CHG_USR_ID, notfcn_id",
            'Line 764': f"SELECT TRUNC(NOTFCN_CRTE_TMS), TRUNC(LST_NOTFCN_TMS), notfcn_id, NOTFCN_STAT_TYP, COUNT(*) FROM gc_owner.ft_T_NTEL WHERE MSG_TYP IN ('ASX_Security_Masters', 'BBEquityCollateral', 'BBGlobalEquity', 'BBGlobalEquity Pricing', 'BBGlobalMutualFund', 'BBGlobalMutual Fund Pricing', 'BBGlobalWarrants', 'BBGlobalWarrantsPricing', 'BBShortSellingRestriction', 'CGS', 'RTDSEPlus_Equity_REF_CSTM', 'RTDSEPlus_EquityWarrants_REF_CSTM', 'RTDSEPlus_Mutual Fund&&MM_REF_CSTM', 'SwapsTradingHours', 'TR_DS_FATCA_Bulk_Solution', 'TR_DS_FATCA_Bulk_Solution_DELTA', 'DBTKFSD', 'DBTKFAD', 'Lipper_Holding', 'Lipper_Fund','D WS_Holding', 'DWS_Fund','FINFRAG_TRADE') AND SOURCE_ID IN ('GC_OWNER@LNORS1P_APP.UK.DB.COM','TRANSLATION','PRDSReconciliation') AND TRUNC(NOTFCN_CRTE_TMS) >= TO_DATE(:prev_month_start, 'DD-MON-YY') AND TRUNC(NOTFCN_CRTE_TMS) < TO_DATE(:curr_month_start, 'DD-MON-YY') AND NOTFCN_STAT_TYP IN ('OPEN','ASSIGN') AND NOTFCN_ID NOT IN ('2','5','23','527','541','640','655','665','667','50001','50009','50016','50021','50050','50052','60202','60981','70006','70007','70085','70086','70087','16','67501','153') GROUP BY TRUNC(NOTFCN_CRTE_TMS), TRUNC(LST_NOTFCN_TMS), notfcn_id, NOTFCN_STAT_TYP"
        }


        # Execute queries and save data to Excel
        execute_queries_and_save_to_excel(connection, queries, file_path, first_day_prev_month_str, first_day_current_month_str)

    except cx_Oracle.Error as error:
        print("Oracle Database Error:", error)
    finally:
        if connection:
            connection.close()
            print("Connection closed.")

username = ""
pwd = ""
dbname = ""
excel_folder_path = ""


if __name__ == '__main__' :
    print('Initializing connection')
    test_connection = db_connection(username, pwd, dbname, excel_folder_path)
