import pandas as pd
import numpy as np
import openpyxl

def read_table(file_name: str, table_name: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(file_name, read_only= False, data_only = True) # openpyxl does not have table info if read_only is True; data_only means any functions will pull the last saved value instead of the formula
    for sheetname in wb.sheetnames: # pulls as strings
        sheet = wb[sheetname] # get the sheet object instead of string
        if table_name in sheet.tables: # tables are stored within sheets, not within the workbook, although table names are unique in a workbook
            tbl = sheet.tables[table_name] # get table object instead of string
            tbl_range = tbl.ref #something like 'C4:F9'
            break # we've got our table, bail from for-loop
    data = sheet[tbl_range] # returns a tuple that contains rows, where each row is a tuple containing cells
    content = [[cell.value for cell in row] for row in data] # loop through those row/cell tuples
    header = content[0] # first row is column headers
    rest = content[1:] # every row that isn't the first is data
    df = pd.DataFrame(rest, columns = header)
    wb.close()
    return df


df = read_table('C:\\Users\\katikako\\Desktop\\HELIOS db.xlsx', 'HeliosDatabase')
#print(df)
table = df.pivot_table(index='SEX', values='AGE',aggfunc=np.count_nonzero)
print(table)