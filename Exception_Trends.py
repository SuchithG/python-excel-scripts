import pandas as pd

input_file_path = 'path_to_your_input_excel_file.xlsx'

try:
    # Define the columns to read
    columns_to_read = ["MSG_TYP", "PRIORITY", "STATUS"]

    # Read specific columns from the sheets
    q1_deal = pd.read_excel(input_file_path, sheet_name='Q1 Deal', usecols=columns_to_read)
    q1_tranche = pd.read_excel(input_file_path, sheet_name='Q1 Tranche', usecols=columns_to_read)

    # Concatenate the data from both sheets
    combined_data = pd.concat([q1_deal, q1_tranche])

    # Create the output DataFrame
    output_df = combined_data.copy()

    # Check if the columns exist before assigning them
    if all(col in combined_data.columns for col in ["MSG_TYP", "PRIORITY", "STATUS"]):
        output_df["MSG_TYP"] = combined_data["MSG_TYP"]
        output_df["PRIORITY"] = combined_data["PRIORITY"]
        output_df["STATUS"] = combined_data["STATUS"]
    else:
        print("Error: One or more columns not found in combined data.")
        

    # Write to output file
    output_file_path = 'path_to_your_output_excel_file.xlsx'
    output_df.to_excel(output_file_path, index=False)

    print("Data processed and written to output file.")

except Exception as e:
    print(f"An error occurred: {e}")
