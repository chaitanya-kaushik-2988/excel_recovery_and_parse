from xlwt import Workbook
import pandas as pd

filename = 'forParsing_task.xls'

# Open the file and read its contents
with open(filename, 'r') as file:
    data = file.readlines()

xldoc = Workbook()
sheet = xldoc.add_sheet("Sheet1", cell_overwrite_ok=True)

for i, row in enumerate(data):
    for j, val in enumerate(row.replace('\n', '').split('\t')):
        sheet.write(i, j, val)
    
xldoc.save('myexcel.xls')

# Define the values that indicate rows to be excluded
exclude_values = ['Customer', 'Code', 'Name', 'City']

# Read the Excel file without header
df = pd.read_excel("myexcel.xls", header=None)

# Filter out rows that contain the exclude values
df = df[~df.apply(lambda row: any(val in str(cell) for cell in row for val in exclude_values), axis=1)]

# Filter out rows that have multiple "-" as content
df = df[~df.apply(lambda row: row.str.count('-').sum() > 1, axis=1)]

# Exclude rows that contain only empty cells
df = df.dropna(how='all')

# Define the header row to be excluded
header_row = "    |      Stat|Account |No|Date |Net due dt|     ..."

# Remove all occurrences of the header row from the DataFrame
df = df[df.apply(lambda row: header_row not in str(row), axis=1)]

# Split the single column into multiple columns based on the delimiter "|",
# trimming whitespace from each cell
filtered_df = df[0].str.strip().str.strip("|").str.split("|", expand=True)

# Define the column names directly
header_columns = ["Stat", "Account", "No", "Date", "Net due dt", "LC amnt", "DD", "CCAr", "PayT", "Type"]

# Set the column names
filtered_df.columns = header_columns

# Save the filtered DataFrame to a CSV file
filtered_df.to_csv("filtered_data.csv", index=False)

print("Filtered data saved to filtered_data.csv")
