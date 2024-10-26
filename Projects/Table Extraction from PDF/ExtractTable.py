import pdfplumber
import os
import pandas as pd
from datetime import datetime

folder_path = os.getcwd() + "\\Projects\\Table Extraction from PDF\\"
excel_file = os.path.join(folder_path, "output.xlsx")
data = []

#loop files for pdf files
for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        file_path = os.path.join(folder_path, filename)
        
        with pdfplumber.open(file_path) as pdf:
            #data from last page only are needed
            last_page = pdf.pages[-1]
            text = last_page.extract_text()
            #print(text) #debug
            
            # Check for "II. KEY STATISTICS"
            if "II. KEY STATISTICS" in text:
                table = last_page.extract_table()
                #print(table) #debug
                
                #table found
                if table:
                    column_names = table[0]  #first row is column headers
                    last_row_data = table[-1] #data found in last row - assumed sums of column data
                    date_added = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    #get each column and their data
                    for col_name, col_data in zip(column_names, last_row_data):
                        data.append([filename, date_added, col_name, col_data])

df = pd.DataFrame(data, columns=["Filename", "Date added", "Column", "Amount"])

#add dataframe to excel file - output.xlsx
with pd.ExcelWriter(excel_file, mode='w') as writer:
    df.to_excel(writer, index=False, sheet_name="Data")
