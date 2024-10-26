import pandas as pd
import os
import matplotlib.pyplot as plt

"""""
#select the file for the report
from tkinter import Tk
from tkinter.filedialog import askopenfilename

Tk().withdraw()

excel_file = askopenfilename(title="Select an Excel file", filetypes=[("Xlsx files", "*.xlsx")])

#current_path = os.getcwd() + "\\Projects\\Data Science training\\"
#excel_file = os.path.join(current_path, "Sessions.xlsx")

if excel_file:
    df = pd.read_excel(excel_file, sheet_name="SessionsData")

    session_count = len(df)
    unique_id_count = df["ID"].nunique()
    unique_family_ids_count = df["Family ID"].nunique()
    unique_family_ids = df.drop_duplicates(subset='Family ID') #keeps only first entry
    total_unique_size = unique_family_ids['Size'].sum()
    family_id_max_size_only = df.loc[df.groupby('Family ID')['Size'].idxmax()]
    family_id_max_size = family_id_max_size_only['Size'].sum()

    print(f"Total sessions: {session_count}")
    print(f"Unique IDs: {unique_id_count}")
    print(f"Unique Family Ds: {unique_family_ids_count}")
    print(f"Correct sum of Family members: {total_unique_size}")
    print(f"Correct sum of Family members with highest Size per Family: {family_id_max_size}")
else:
    print("No file selected.")
"""""

current_path = os.getcwd() + "\\Projects\\Data Science training\\"
excel_file = os.path.join(current_path, "Sessions.xlsx")
df = pd.read_excel(excel_file, sheet_name="SessionsData")

session_count = len(df)
unique_id_count = df["ID"].nunique()
unique_family_ids_count = df["Family ID"].nunique()
unique_family_ids = df.drop_duplicates(subset='Family ID') #keeps only first entry
total_unique_size = unique_family_ids['Size'].sum()
family_id_max_size_only = df.loc[df.groupby('Family ID')['Size'].idxmax()]
family_id_max_size = family_id_max_size_only['Size'].sum()

size_family_id_count = df.groupby('Size')['Family ID'].nunique().reset_index(name='Family ID Count')


print(f"Total sessions: {session_count}")
print(f"Unique IDs: {unique_id_count}")
print(f"Unique Family Ds: {unique_family_ids_count}")
print(f"Correct sum of Family members: {total_unique_size}")
print(f"Correct sum of Family members with highest Size per Family: {family_id_max_size}")


#line_chart
plt.figure(figsize=(10, 6))
plt.plot(size_family_id_count['Size'], size_family_id_count['Family ID Count'], marker='o', label='Family ID Count')
#plt.plot(size_family_id_count['Family ID Count'], size_family_id_count['Size'], marker='o', label='Family ID Count')

#data labels - must be after plot!
for i, row in size_family_id_count.iterrows():
        plt.text(row['Size'], row['Family ID Count'], str(row['Family ID Count']), 
                 ha='center', va='bottom') #alignment

plt.title('Size by Family ID')
plt.xlabel('Family ID')
plt.ylabel('Size')
#plt.xticks(rotation=30)
plt.grid()
plt.tight_layout() #fit
plt.show()