#In this I am trying to take out the data from excel file

import openpyxl

# Load the Excel workbook
workbook = openpyxl.load_workbook('Data.xlsx')

# Select the sheet containing the table
sheet = workbook['genl2_Table']

print(sheet)

# Find the first cell in the table (assuming it starts at A1)
table_start = 'A1'
while sheet[table_start].value is None:
    # Move to the next cell until a non-empty cell is found
    table_start = sheet[table_start].offset(row=1, column=0).coordinate

# Find the last cell in the table
table_end = table_start
while sheet[table_end].value is not None:
    # Move to the next cell until an empty cell is found
    table_end = sheet[table_end].offset(row=0, column=1).coordinate

# Convert the table data into an array of arrays
table_data = []
for row in sheet[table_start:table_end]:
    row_data = []
    for cell in row:
        row_data.append(cell.value)
    table_data.append(row_data)

# Print the table data
print(table_data)






