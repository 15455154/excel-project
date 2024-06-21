import pandas as pd

# Correct file path
file_path = r"C:\Users\prana\project.txt"

# Read and process the file
with open(file_path, 'r') as file:
    lines = file.readlines()

# Create a DataFrame from the file contents (assuming comma-separated values)
data = [line.strip().split(',') for line in lines]

# Check if the first row is a header or not
header = data[0]
num_columns = len(header)
columns = [f'Column{i}' for i in range(1, num_columns + 1)] if len(data) > 1 else []

# Create DataFrame with or without headers
if columns:
    df = pd.DataFrame(data[1:], columns=header)
else:
    df = pd.DataFrame(data, columns=columns)

# Write the DataFrame to an Excel file
output_excel_path = 'output.xlsx'
df.to_excel(output_excel_path, index=False)

# Optionally, write to a new Excel file with a specific sheet name
output_excel_path_with_sheet = 'my_data.xlsx'
df.to_excel(output_excel_path_with_sheet, sheet_name='MySheet', index=False)

print(f"Data has been successfully written to {output_excel_path} and {output_excel_path_with_sheet}")
