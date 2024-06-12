# Define the input and output file paths for the new uploaded file
input_file_path = '/mnt/data/file-gnJHC1LM6bC1dtwK2ZbNLnoa'
output_file_path = '/mnt/data/cleaned_data_v.txt'

# Read the content of the input file
with open(input_file_path, 'r', encoding='utf-8', errors='ignore') as file:
    lines = file.readlines()

# Initialize an empty list to hold the cleaned lines
cleaned_lines = []
header_found = False

# Define keywords and phrases to remove
keywords_to_remove = [
    "PRICE UPDATED",
    "PRICE ALREADY EXISTS",
    "PRICE SUPPLIED FOR A SECURITY WHICH IS NOT IN SELECTION FILE",
    "***",
    "---",
    "No. OF SECURITIES PROCESSED:",
    "No. OF SECURITIES NOT UPDATED:",
    "No. OF SECURITIES UPDATED:"
]

# Define headers
header_line = "SEC NO, SHORT NAME, TK NO, EXCH, CURR, DATE, PRICE, REMARKS"

# Iterate over the lines to clean them
for line in lines:
    # Skip lines with specified keywords
    if any(keyword in line for keyword in keywords_to_remove):
        continue
    
    # Ensure headers are present only in the first row
    if header_line in line:
        if header_found:
            continue
        header_found = True
    
    cleaned_lines.append(line)

# Write the cleaned data to the output file
with open(output_file_path, 'w', encoding='utf-8') as file:
    for line in cleaned_lines:
        file.write(line)

print("Data cleaning complete. The cleaned data is saved in 'cleaned_data_v.txt'.")
