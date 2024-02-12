# Excel File Recovery and Parsing Script

This Python script is designed to recover data from a potentially corrupt Excel (.xls) file and parse it to create a sanitized CSV file. It performs the following tasks:

1. **Recovery of Data**: Utilizes the `xlwt` library to recover data from the corrupt Excel file.
2. **Exclusion of Specified Rows**: Excludes rows based on predefined values such as 'Customer', 'Code', 'Name', and 'City'.
3. **Filtering Rows with Multiple Dashes**: Removes rows containing multiple dashes ('-') as content.
4. **Exclusion of Empty Rows**: Eliminates rows that contain only empty cells.
5. **Exclusion of Header Row**: Removes occurrences of a predefined header row.
6. **Splitting Columns and Cleaning Data**: Splits a single column into multiple columns based on the delimiter "|", trimming whitespace from each cell.
7. **Setting Column Names**: Sets column names for the parsed data.
8. **Saving to CSV**: Saves the sanitized data to a CSV file.

## Python Version
- Python 3.8.10

## Usage

1. **Install Dependencies**: Ensure you have the required dependencies installed from requirements.txt. You can install them using pip:

2. **Run the Script**: Execute the Python script `recover_and_parse.py`. Make sure to replace `'forParsing_task.xls'` with the path to your corrupt Excel file.

3. **Review Output**: Once the script completes, review the generated `filtered_data.csv` file, which contains the sanitized data.

## Note

- Make sure to replace the file paths and header row content with your specific requirements.
- This script assumes that the corrupt Excel file is in the `.xls` format.
