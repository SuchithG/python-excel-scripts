import pandas as pd

def process_txt_to_excel(input_file, output_file):
    # Read the data from the text file
    with open(input_file, 'r') as file:
        lines = file.readlines()
    
    # Process the lines to remove those starting with '*' or '-'
    processed_lines = [line for line in lines if not (line.startswith('*') or line.startswith('-'))]
    
    # Split each line into columns (assuming space-separated values)
    data = [line.split() for line in processed_lines]
    
    # Create a DataFrame from the processed data
    df = pd.DataFrame(data)
    
    # Save the DataFrame to an Excel file
    df.to_excel(output_file, index=False, header=False)

# Example usage
input_file = 'path_to_input_file.txt'
output_file = 'output_file.xlsx'
process_txt_to_excel(input_file, output_file)
