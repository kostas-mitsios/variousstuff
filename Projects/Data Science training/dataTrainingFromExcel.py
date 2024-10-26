import pandas as pd
import os

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

print(f"Total sessions: {session_count}")
print(f"Unique IDs: {unique_id_count}")
print(f"Unique Family Ds: {unique_family_ids_count}")
print(f"Correct sum of Family members: {total_unique_size}")
print(f"Correct sum of Family members with highest Size per Family: {family_id_max_size}")