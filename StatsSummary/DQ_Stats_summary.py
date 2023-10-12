import pandas as pd
import time

start_time = time.time()

def adjust_month_format(value):
    if pd.isna(value):
        return value
    if "_" in str(value):
        return str(value).replace("_", "-")
    try:
        date_val = pd.to_datetime(value)
        return date_val.strftime('%b-%y')
    except:
        return value

def adjust_date_formats(df):
    df['Date'] = df['Date'].dt.strftime('%m/%d/%Y')
    df['Month'] = df['Month'].apply(adjust_month_format)
    df['Actual Date of upload'] = df['Actual Date of upload'].dt.strftime('%Y-%m-%d')
    return df

reference_file_path = "path/to/reference/excelfile.xlsx"
try:
    reference_df = pd.read_excel(reference_file_path, engine='openpyxl')
    print("Successfully loaded the reference file.")
except Exception as e:
    print(f"An error occurred while loading the reference file: {e}")
    exit()

excel_file_path = "path/to/your/main/excelfile.xlsx"
try:
    all_sheets = pd.read_excel(excel_file_path, sheet_name=None, engine='openpyxl')
    print("Successfully loaded the main Excel workbook.")
except Exception as e:
    print(f"An error occurred while loading the main Excel workbook: {e}")
    exit()

output_path = "path/to/your/desired_output_file.xlsx"

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    for sheet_name, df in all_sheets.items():
        
        df = adjust_date_formats(df)
        
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Successfully saved all updated sheets to the output Excel workbook at {output_path}.")
end_time = time.time()

execution_time_minutes = (end_time - start_time) / 60
print(f"Script execution in {execution_time_minutes:.2f} minutes")
