import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

def transform_data(input_directory, output_file_path):
    # Get the previous month and year
    last_month = datetime.now() - relativedelta(months=1)
    previous_month = last_month.strftime('%b').upper()  # Abbreviated and uppercased month (e.g., 'OCT')
    previous_year = last_month.strftime('%Y')  # Full year (e.g., '2023')

    # Construct the input file name based on the previous month and year
    input_file_name = f'ExceptionTrends_{previous_month}_{previous_year}_script_output.xlsx'
    input_file_path = f'{input_directory}/{input_file_name}'

    # Load the input data
    data = pd.read_excel(input_file_path)

    # Group by 'Exception trend', 'MSG_TYP', 'PRIORITY', 'STATUS' and sum the 'Volume'
    grouped_data = data.groupby(['Exception trend', 'MSG_TYP', 'PRIORITY', 'STATUS']).agg(Total_Volume=('Volume', 'sum')).reset_index()

    # Pivot the data to create a multi-level column structure for each 'PRIORITY' and 'STATUS'
    pivot_data = grouped_data.pivot_table(index=['Exception trend', 'MSG_TYP'], columns=['PRIORITY', 'STATUS'], values='Total_Volume', fill_value=0)

    # Flatten the multi-level columns
    pivot_data.columns = [' '.join(col).strip() for col in pivot_data.columns.values]

    # Calculate the total and closed volumes for each priority
    for priority in ['P1', 'P2', 'P3']:
        pivot_data[f'Total {priority}'] = pivot_data[f'{priority} CLOSED'] + pivot_data[f'{priority} OPEN']
        pivot_data[f'Closed {priority}'] = pivot_data[f'{priority} CLOSED']

    # Calculate 'Total OPEN' and 'OPEN Percentage'
    pivot_data['Total OPEN'] = pivot_data[['P1 OPEN', 'P2 OPEN', 'P3 OPEN']].sum(axis=1)
    pivot_data['Total Volume'] = pivot_data[['Total P1', 'Total P2', 'Total P3']].sum(axis=1)
    pivot_data['OPEN Percentage'] = (pivot_data['Total OPEN'] / pivot_data['Total Volume']).fillna(0) * 100

    # Add the 'Month/Year' column to the dataframe, formatted as 'Oct-2023'
    pivot_data['Month/Year'] = last_month.strftime('%b-%Y')

    # Filter out rows where 'MSG_TYP' has all zero values in the specified columns
    filter_columns = ['P1 OPEN', 'P2 OPEN', 'P3 OPEN', 'Closed P1', 'Closed P2', 'Closed P3', 'Total OPEN']
    pivot_data = pivot_data[~(pivot_data[filter_columns] == 0).all(axis=1)]

    # Reset index to turn the MultiIndex into columns
    final_output = pivot_data.reset_index()

    # Reorder and select columns for the final output
    final_output_columns = ['Month/Year', 'Exception trend', 'MSG_TYP', 'Total P1', 'Total P2', 'Total P3', 'Closed P1', 'Closed P2', 'Closed P3', 'Total OPEN', 'OPEN Percentage']
    final_output = final_output[final_output_columns]

    # Save the transformed data to an Excel file
    final_output.to_excel(output_file_path, index=False)
    return "Data transformation complete. Output saved to: " + output_file_path

# Example usage of the script
input_directory = '/path/to/your/input/directory'
output_file_path = '/path/to/your/output/file.xlsx'
print(transform_data(input_directory, output_file_path))
