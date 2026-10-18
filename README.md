# Excel to CSV Data Processing
## Project Overview
This project demonstrates how to read, process, and export employee data from an Excel file (.xlsx)
to a CSV file (.csv) using Python. It uses the Pandas and OpenPyXL libraries to perform different file-
handling operations.

## Key Objectives
- Read an Excel file using two methods: Pandas and OpenPyXL.
- Rename DataFrame columns using a dictionary.
-Check whether specific columns exist in a DataFrame.
- Display and inspect employee data.
-Convert column names to uppercase and replace underscores with spaces.
-Export DataFrame data to a CSV file using Pandas.
-Export Excel worksheet data to a CSV file using Python's built-in 'csv module and OpenPyXL.

## Libraries Used
**Pandas:** For reading Excel files, manipulating DataFrames, renaming columns, and exporting data
to CSV.
-**OpenPyXL:** For reading Excel worksheets and extracting their data.
-**CSV:** For writing data into CSV files.

## Functions Description
###1. 'data_reader()
Reads an Excel file using Pandas and returns its contents as a DataFrame.
### 2. 'data_reader1()
Reads the 'Trainee_Data worksheet using OpenPyXL and returns the data as a list of lists.
### 3. 'rename_columns()
Renames specified DataFrame columns according to a dictionary. It also checks whether the original
column names exist.
### 4. 'check_columns_exist().
Checks whether one or more specified columns exist in a DataFrame and returns a dictionary
containing Boolean results.
### 5. 'write_to_csv()
Exports a Pandas DataFrame to a CSV file without including the DataFrame index.
### 6. 'write_to_csv1(
Reads data from an Excel worksheet using OpenPyXL and writes it to a CSV file using Python's built-in
'csv module.

## Conclusion
This project provides practical experience in reading Excel files, manipulating tabular data, validating
column names, and exporting data to CSV using different Python libraries.
