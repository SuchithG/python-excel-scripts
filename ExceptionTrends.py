import pandas as pd

# Load the data from Excel
file_path = 'path_to_your_excel_file.xlsx'  # Update this with the path to your Excel file
df = pd.read_excel(file_path, sheet_name='Q1 Deal')

# Ensure the EFFECTIVE_DTE column is in datetime format
df['CLOSE_DTE'] = pd.to_datetime(df['CLOSE_DTE'])

# Extract Month/Year from EFFECTIVE_DTE
df['Month/Year'] = df['CLOSE_DTE'].dt.strftime('%b-%y')

# Group by MSG_TYP, PRIORITY, STATUS, and Month/Year
grouped = df.groupby(['MSG_TYP', 'PRIORITY', 'STATUS', 'Month/Year'])

# Count STATUS and MSG_TYP
output_df = grouped.agg(Status_Count=('STATUS', 'count'), Volume=('MSG_TYP', 'count')).reset_index()

print(output_df)
