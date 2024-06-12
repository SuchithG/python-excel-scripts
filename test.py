# Define the input and output file paths for the new uploaded file
input_file_path = '/mnt/data/file-JeTdGV5V6eRWGW6mzYCoTkQK'
output_file_path = '/mnt/data/cleaned_data_v8.txt'

# Read the content of the input file
with open(input_file_path, 'r', encoding='utf-8', errors='ignore') as file:
    lines = file.readlines()

# Initialize an empty list to hold the cleaned lines
cleaned_lines = []

# Iterate over the lines to clean them
for line in lines:
    if lines.startswith(('*', '-', '^')):
        continue
    
    if any(keyword in line for keyword in ["PRICE UPDATED", "PRICE ALREADY EXISTS", "PRICE SUPPLIED FOR A SECURITY WHICH IS NOT IN SELECTION FILE"]):
        continue
    cleaned_lines.append(line)

# Write the cleaned data to the output file
with open(output_file_path, 'w', encoding='utf-8') as file:
    for line in cleaned_lines:
        file.write(line)

print("Data cleaning complete. The cleaned data is saved in 'cleaned_data_v8.txt'.")
