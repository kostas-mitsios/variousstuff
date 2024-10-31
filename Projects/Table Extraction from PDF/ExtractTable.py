import pdfplumber
import os
import pandas as pd
from datetime import datetime
import re
import tkinter as tk
from tkinter import filedialog

#folder selector
def select_folder():
    root = tk.Tk()
    #hide main window
    root.withdraw()
    folder_selected = filedialog.askdirectory(title="Select Folder containing PDF files")
    return folder_selected

#prompt the user for the folder
folder_path = select_folder()
if not folder_path:
    print("No folder selected, exiting.") #bruh
    exit()

#generate excel file name. Name will be MergedResult_yyyymmddhhmm
timestamp = datetime.now().strftime("%Y%m%d%H%M")
excel_file = os.path.join(folder_path, f"MergedResult_{timestamp}.xlsx")
data = []
failed_files = []

#regex to match table titles
pattern = r"(\d{2}-\d{2}-\d{4})\s+(\d{2}:\d{2})\s+(\d+)\s+(\d+)\s+(.+)"
relevant_titles = ["Τεχνικός Ασφάλειας:", "Iατρός Εργασίας:"]

#loop folder for pdf files
for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        file_path = os.path.join(folder_path, filename)
        
        with pdfplumber.open(file_path) as pdf:
            found_relevant_data = False
            
            #loop pages
            for page in pdf.pages:
                text = page.extract_text()

                #check for table title
                if "Επισκέψεις Τεχνικών" in text or "Επισκέψεις Ιατρών" in text:
                    found_relevant_data = True
                    person_column = ""  #initialize occupation column
                    
                    #select appropriate column to store name
                    if "Επισκέψεις Τεχνικών" in text:
                        person_column = "Τεχνικός Ασφάλειας:"
                    elif "Επισκέψεις Ιατρών" in text:
                        person_column = "Iατρός Εργασίας:"
                    
                    #splitter of lines is '\n'
                    lines = text.split('\n')
                    start_index = None
                    for i, line in enumerate(lines):
                        if any(title in line for title in relevant_titles):
                            start_index = i + 1
                            break
                    
                    #start_index not null => table title was found
                    if start_index is not None:
                        for line in lines[start_index:]:
                            if "Αποθηκευμένο αρχείο" in line:
                                break
                            
                            match = re.match(pattern, line)
                            if match:
                                date, time, hours, minutes, person = match.groups()
                                date_added = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                                #initialize variables as one will have to be null
                                technician = ""
                                doctor = ""
                                
                                if person_column == "Τεχνικός Ασφάλειας:":
                                    technician = person
                                elif person_column == "Iατρός Εργασίας:":
                                    doctor = person
                                
                                #append list column data
                                data.append([
                                    filename,
                                    date_added,
                                    str(date),
                                    str(time),
                                    int(hours),
                                    int(minutes),
                                    technician,
                                    doctor
                                ])
                #eof
                if "Αποθηκευμένο αρχείο" in text:
                    break
            #no table title that we want was found. File will be added to txt file for review
            if not found_relevant_data:
                failed_files.append(filename)  

columns = ["Filename", "Date added", "Ημερομηνία", "Ώρα", "[Διάρκεια] Ώρες", "Λεπτά", "Τεχνικός Ασφάλειας:", "Iατρός Εργασίας:"]
df = pd.DataFrame(data, columns=columns)

#read file if exists and add data to dataframe so that we can append the next ones and save back
if os.path.exists(excel_file):
    existing_df = pd.read_excel(excel_file, sheet_name="Data")
    df = pd.concat([existing_df, df], ignore_index=True)

#save data in excel
with pd.ExcelWriter(excel_file, mode='w') as writer:
    df.to_excel(writer, index=False, sheet_name="Data")

#failures - utf-8 must be selected
if failed_files:
    with open(os.path.join(folder_path, "failed_to_open.txt"), 'w', encoding='utf-8') as f:
        for failed_file in failed_files:
            f.write(f"{failed_file}\n")
