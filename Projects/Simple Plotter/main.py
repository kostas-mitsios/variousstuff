import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
import win32com.client as win32
import win32api
import win32gui
import time
import win32com

def select_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            # Load workbook to get sheet names
            workbook = pd.ExcelFile(file_path)
            
            # Load and print Power Queries
            load_power_queries(file_path)
            
            # Show sheet selection window
            show_sheet_selection(workbook)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open Excel file: {e}")

def load_power_queries(file_path):
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  # Keep Excel hidden in the background
        workbook = excel.Workbooks.Open(file_path)

        # Retrieve Power Queries
        queries = workbook.Queries
        if queries.Count > 0:
            print("Power Queries found in the workbook:")
            for query in queries:
                print(f" - {query.Name}")
        else:
            print("No Power Queries found in this workbook.")

        #load_and_refresh_specific_query(file_path, "IndividualsEnrolled")
        load_and_refresh_specific_connection(file_path, "Query - IndividualsEnrolled")
        #workbook.Close(SaveChanges=False)
        #excel.Quit()
        #full name in VBA is "Query - IndividualsEnrolled"
        """""
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(file_path)
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        xlapp.Quit()
        """""

    except Exception as e:
        print(f"Error loading Power Queries: {e}")

def load_and_refresh_specific_query(file_path, query_name):
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True  # Set to True if you want to see the refresh in action
        workbook = excel.Workbooks.Open(file_path)

        # Retrieve Power Queries
        queries = workbook.Queries
        query_names = [query.Name for query in queries]

        
        if query_name in query_names:
            # Find the specific query and refresh it
            specific_query = next(query for query in queries if query.Name == query_name)
            print(f"Refreshing query: {specific_query.Name}")
            print(specific_query.__class__)
            specific_query.Refresh()  # Refresh the specific query
            # Wait for the query to finish refreshing
            while specific_query.Refreshing:
                print("Query is still refreshing...")
                time.sleep(1)  # Sleep for a short time to avoid maxing out the CPU
            print("Query refresh complete.")
        else:
            print(f"Query '{query_name}' not found in the workbook. Available queries are: {query_names}")
        


        workbook.Close(SaveChanges=True)
        excel.Quit()

    except Exception as e:
        print(f"Error refreshing specific query: {e}")

def load_and_refresh_specific_connection(file_path, connection_name):
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True  # Set to True if you want to see the refresh in action
        workbook = excel.Workbooks.Open(file_path)

        # Retrieve workbook connections
        connections = workbook.Connections
        connection_names = [connection.Name for connection in connections]

        if connection_name in connection_names:
            # Find the specific connection and refresh it
            specific_connection = next(connection for connection in connections if connection.Name == connection_name)
            print(f"Refreshing connection: {specific_connection.Name}")
            print(specific_connection.__class__)
            specific_connection.Refresh()  # Refresh the specific connection

            # Wait for the connection to finish refreshing
            while specific_connection.Refreshing:
                print("Connection is still refreshing...")
                time.sleep(1)  # Sleep for a short time to avoid maxing out the CPU
            print("Connection refresh complete.")
        else:
            print(f"Connection '{connection_name}' not found in the workbook. Available connections are: {connection_names}")

        workbook.Close(SaveChanges=True)  # Save changes before closing
        excel.Quit()

    except Exception as e:
        print(f"Error refreshing specific connection: {e}")


def show_sheet_selection(workbook):
    sheet_window = tk.Toplevel(root)
    sheet_window.title("Select a Sheet")
    sheet_window.geometry("300x200")

    sheet_names = workbook.sheet_names
    num_sheets = len(sheet_names)
    num_cols = 2
    num_rows = (num_sheets + 1) // num_cols

    frame = tk.Frame(sheet_window)
    frame.grid(padx=20, pady=20)

    for idx, sheet_name in enumerate(sheet_names):
        button = tk.Button(
            frame, text=sheet_name, 
            command=lambda name=sheet_name: select_sheet(workbook, name)
        )
        row, col = divmod(idx, num_cols)
        button.grid(row=row, column=col, padx=5, pady=5)

    sheet_window.update_idletasks()
    min_width = max(300, sheet_window.winfo_reqwidth() + 40)
    min_height = max(200, sheet_window.winfo_reqheight() + 40)
    sheet_window.geometry(f"{min_width}x{min_height}")

def select_sheet(workbook, sheet_name):
    try:
        df = pd.read_excel(workbook, sheet_name=sheet_name, engine="openpyxl")
        tables = find_tables_in_sheet(df, sheet_name)
        if len(tables) == 1:
            table_name = tables[0]
            select_table(df, sheet_name, table_name)
        elif len(tables) > 1:
            show_table_selection(df, tables, sheet_name)
        else:
            messagebox.showwarning("Warning", "No tables found. //TODO")
    except Exception as e:
        messagebox.showerror("Error", f"Could not read sheet: {e}")

def find_tables_in_sheet(df, sheet_name):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]
    return [table.name for table in sheet.tables.values()]

def show_table_selection(df, tables, sheet_name):
    table_window = tk.Toplevel(root)
    table_window.title("Select a Table")
    table_window.geometry("300x200")

    num_tables = len(tables)
    num_cols = 2
    num_rows = (num_tables + 1) // num_cols

    frame = tk.Frame(table_window)
    frame.grid(padx=20, pady=20)

    for idx, table_name in enumerate(tables):
        button = tk.Button(
            frame, text=table_name, 
            command=lambda t_name=table_name: select_table(df, sheet_name, t_name)
        )
        row, col = divmod(idx, num_cols)
        button.grid(row=row, column=col, padx=5, pady=5)

    table_window.update_idletasks()
    min_width = max(300, table_window.winfo_reqwidth() + 40)
    min_height = max(200, table_window.winfo_reqheight() + 40)
    table_window.geometry(f"{min_width}x{min_height}")

def select_table(df, sheet_name, table_name):
    headers = list(df.columns)
    if headers:
        open_excel_and_move_cursor(sheet_name, "A2")  # Set to "A2" as the target cell
    else:
        messagebox.showwarning("Warning", "No headers found. //TODO")

def open_excel_and_move_cursor(sheet_name, cell_address):
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        workbook = excel.Workbooks.Open(file_path)
        sheet = workbook.Sheets(sheet_name)
        sheet.Activate()

        # Select the specific cell (e.g., A2)
        cell = sheet.Range(cell_address)
        cell.Select()

        # Get Excel window handle and position
        excel_hwnd = win32gui.FindWindow(None, excel.Caption)
        rect = win32gui.GetWindowRect(excel_hwnd)
        excel_left, excel_top = rect[0], rect[1]

        # Calculate screen coordinates of the cell within the window
        x_offset = cell.Left
        y_offset = cell.Top
        screen_x = int(excel_left + x_offset)
        screen_y = int(excel_top + y_offset)

        print(f"Cursor screen coordinates: x={screen_x}, y={screen_y}")
        win32api.SetCursorPos((screen_x, screen_y))

    except Exception as e:
        messagebox.showerror("Error", f"Could not open Excel application or move cursor: {e}")

# Main window
root = tk.Tk()
root.title("Excel Table Selector")
root.geometry("300x200")

open_file_btn = tk.Button(root, text="Select Excel File", command=select_file)
open_file_btn.pack(expand=True)

root.mainloop()
