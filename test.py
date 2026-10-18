import csv
from typing import Dict, List, Union
import openpyxl
import pandas as pd

file_path = '/home/leenmhmod30/estarta/trainee_excel_csv_exercise.xlsx'
csv_file_path = '/home/leenmhmod30/estarta/trainee_excel_csv_exercise.csv'

#read excel file using pandas
def data_reader(file_path: str = file_path) -> pd.DataFrame:

    """Reads an Excel file using pandas and returns its content as a DataFrame."""
    df = pd.read_excel(file_path)

    return df

#read excel file using openpyxl

def data_reader1(file_path: str = file_path) -> List[List]:

    """Reads an Excel file using openpyxl and returns its content as a list of lists."""

    wb = openpyxl.load_workbook(file_path)
    ws = wb["Trainee_Data"]
    data = [list(row) for row in ws.iter_rows(values_only=True)]    
    return data


#rename columns in the dataframe

def rename_columns(df: pd.DataFrame, rename_dict: dict) -> pd.DataFrame:

    """Renames columns in the DataFrame based on the provided dictionary."""

    df_columns = df.columns.tolist()

    for old_col in rename_dict.keys():
        if old_col not in df_columns:
            print(f"Error: The column '{old_col}' does not exist!")
            return df
         
    new_df = df.rename(columns=rename_dict)
    return new_df

# check if a column exists in the dataframe
def check_columns_exist( df: pd.DataFrame, columns: Union[str, List[str]]) -> Dict[str, bool]:

    """Checks if the specified columns exist in the DataFrame."""   
    if isinstance(columns, str):
        columns = [columns]

    results = {}
    
    for col in columns:
        if col in df.columns:
            results[col] = True
        else:
            results[col] = False

    return results
#write the dataframe to a csv file    
def write_to_csv(df: pd.DataFrame, output_path: str = csv_file_path):
    """Writes the DataFrame to a CSV file using pandas."""
    
    df.to_csv(output_path, index=False)
   
    return

def write_to_csv1(input_path: str = file_path, output_path: str = csv_file_path):

    """Writes the DataFrame to a CSV file using openpyxl."""

    wb= openpyxl.load_workbook(filename=file_path)
    sheet = wb.active
    csv_data=[]   
    for value in sheet.iter_rows (values_only=True):
        csv_data.append(list(value))
        
    with open(output_path, 'w', newline='') as f:
        writer= csv.writer(f,delimiter=',')
        for line in csv_data:
            writer.writerow(line)


    return


df = data_reader()
print("\nData Reader:")
print(df.head())

print("\nData Reader 1:")
print(data_reader1())

print("\nRenamed Columns:")       
print(rename_columns(df, {
    'Employee_ID': 'Employee ID',
    'Full_Name': 'Full Name',
    'Department': 'Department',
    'Job_Title': 'Job Title',
    'City': 'City',
    'Email': 'Email',
    'Phone_Number': 'Phone Number',
    'Joining_Date': 'Joining Date',
    'Employment_Status': 'Employment Status',
    'Notes': 'Notes'
}))


print(df.columns) 
rename=df.columns.str.replace('_', ' ').str.upper().to_list() 
print(type(rename))
print(rename)
print(check_columns_exist(df, ['Employee_ID', 'Full_Name', 'Department', 'Job_Title', 'City', 'Email', 'Phone_Number', 'Joining_Date', 'Employment_Status', 'Notes']))

write_to_csv(df)
write_to_csv1()