import openpyxl
import json
import os

# Get the current working directory
current_directory = os.path.dirname(os.path.abspath(__file__))
print(current_directory)

# Define the file path relative to the current directory
file_path = os.path.join(current_directory, 'members.xlsx')

# Load the Excel file
workbook = openpyxl.load_workbook(file_path)  # Replace 'members.xlsx' with your file path
sheet = workbook.active

# Convert the sheet data to a list of dictionaries (JSON-like format)
data = []
for row in sheet.iter_rows(min_row=2, values_only=True):  # Assuming the header is in the first row
    record = {}
    record['NAME'] = row[0]
    record['AFFILIATION'] = row[1]
    record['POSITION'] = row[2]
    record['MAJOR'] = row[3]
    record['LINK'] = row[4]
    data.append(record)

# Convert the data list to JSON format
json_data = json.dumps(data, indent=4)
print(json_data)

json_path = os.path.join(current_directory, 'members.json')
# Save the JSON data to a file
with open(json_path, 'w') as file:
    file.write(json_data)
