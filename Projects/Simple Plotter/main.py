import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
#import openpyxl

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            workbook = pd.ExcelFile(file_path)
            show_sheet_selection(workbook)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open Excel file: {e}")

def show_sheet_selection(workbook):
    sheet_window = tk.Toplevel(root)
    sheet_window.title("Select a Sheet")
    sheet_window.geometry("300x200") #starter window size - minimum size

    sheet_names = workbook.sheet_names
    num_sheets = len(sheet_names)
    num_cols = 2 # 2 columns style
    num_rows = (num_sheets + 1) // num_cols

    #frame to center buttons
    frame = tk.Frame(sheet_window)
    frame.grid(padx=20, pady=20)

    for idx, sheet_name in enumerate(sheet_names):
        button = tk.Button(
            frame, text=sheet_name, 
            command=lambda name=sheet_name: select_sheet(workbook, name)
        )
        row, col = divmod(idx, num_cols)
        button.grid(row=row, column=col, padx=5, pady=5) #this is to set rows of buttons

    #dynamically adjust window size
    sheet_window.update_idletasks()
    min_width = max(300, sheet_window.winfo_reqwidth() + 40)
    min_height = max(200, sheet_window.winfo_reqheight() + 40)
    sheet_window.geometry(f"{min_width}x{min_height}")

def select_sheet(workbook, sheet_name):
    #todo - maybe logic needed for single sheet existing
    try:
        df = pd.read_excel(workbook, sheet_name=sheet_name, engine="openpyxl")
        tables = find_tables_in_sheet(df)
        if len(tables) == 1:
            select_table(df, tables[0])
        elif len(tables) > 1:
            show_table_selection(df, tables)
        else:
            messagebox.showwarning("Warning", "No tables found. //TODO")
    except Exception as e:
        messagebox.showerror("Error", f"Could not read sheet: {e}")

def find_tables_in_sheet(df):
    #todo - actually do something if no tables found
    return [df] if not df.empty else []

def show_table_selection(df, tables):
    table_window = tk.Toplevel(root)
    table_window.title("Select a Table")
    table_window.geometry("300x200")  # Set a minimum size

    num_tables = len(tables)
    num_cols = 2
    num_rows = (num_tables + 1) // num_cols

    frame = tk.Frame(table_window)
    frame.grid(padx=20, pady=20)

    for idx, table in enumerate(tables):
        button = tk.Button(
            frame, text=f"Table {idx + 1}", 
            command=lambda t=table: select_table(df, t)
        )
        row, col = divmod(idx, num_cols)
        button.grid(row=row, column=col, padx=5, pady=5)

    table_window.update_idletasks()
    min_width = max(300, table_window.winfo_reqwidth() + 40)
    min_height = max(200, table_window.winfo_reqheight() + 40)
    table_window.geometry(f"{min_width}x{min_height}")

def select_table(df, table):
    headers = list(table.columns)
    if headers:
        show_header_selection(headers)
    else:
        messagebox.showwarning("Warning", "No headers found. //TODO")

def show_header_selection(headers):
    header_window = tk.Toplevel(root)
    header_window.title("Select Headers")
    header_window.geometry("300x200")  # Set a minimum size

    num_headers = len(headers)
    num_cols = 2
    num_rows = (num_headers + 1) // num_cols

    frame = tk.Frame(header_window)
    frame.grid(padx=20, pady=20)

    for idx, header in enumerate(headers):
        button = tk.Button(frame, text=header, command=lambda h=header: select_header(h))
        row, col = divmod(idx, num_cols)
        button.grid(row=row, column=col, padx=5, pady=5)

    header_window.update_idletasks()
    min_width = max(300, header_window.winfo_reqwidth() + 40)
    min_height = max(200, header_window.winfo_reqheight() + 40)
    header_window.geometry(f"{min_width}x{min_height}")

def select_header(header):
    messagebox.showinfo("Selected Header", f"You selected: {header}")

#main window
root = tk.Tk()
root.title("Excel Table Selector")
root.geometry("300x200")

open_file_btn = tk.Button(root, text="Select Excel File", command=select_file)
open_file_btn.pack(expand=True)  #center button

root.mainloop()
