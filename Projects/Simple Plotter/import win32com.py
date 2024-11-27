import win32com.client as win32
import time

def check_if_query_is_running(file_path):
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True  # Optional: set to True if you want to see what's happening
        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.Sheets(1)  # Change to the correct sheet index or name

        # Get the table (ListObject)
        table = worksheet.ListObjects(1)  # Assuming the table is the first table on the sheet. Adjust index as necessary.

        # Store the initial column widths
        initial_column_widths = {}
        for column in range(1, table.Range.Columns.Count + 1):
            column_name = table.Range.Cells(1, column).Value  # Use the first row as the header
            initial_column_widths[column_name] = table.Range.Columns(column).ColumnWidth

        # Retrieve Power Queries
        queries = workbook.Queries
        if queries.Count > 0:
            print("Power Queries found in the workbook:")
            for query in queries:
                print(f" - {query.Name}")
                break
        else:
            print("No Power Queries found in this workbook.")
        
        # Trigger the refresh for the first query (or a specific query)
        query = queries.Item(1)  # Change if necessary to select the specific query
        query.Refresh()
        #excel.CalculateUntilAsyncQueriesDone()
        print("223213")

        # Check if Excel is busy (running background query)
        while not excel.Ready:
            print("Excel is still busy, running background query...")
            time.sleep(1)  # Sleep to avoid maxing out the CPU

        # Wait for column resize detection
        resize_detected = False
        while not resize_detected:
            resize_detected = True
            for column in range(1, table.Range.Columns.Count + 1):
                column_name = table.Range.Cells(1, column).Value  # Use the first row as the header
                current_column_width = table.Range.Columns(column).ColumnWidth
                print(excel.CalculationState)

                if current_column_width != initial_column_widths.get(column_name, None):
                    print(f"Column '{column_name}' width has changed!")
                    resize_detected = True

            if not resize_detected:
                print("No column resize detected yet...")
                time.sleep(1)  # Sleep to avoid busy-waiting

        print("Query refresh complete, and column resize detected.")
        
        # Close the workbook and quit Excel
        workbook.Close(SaveChanges=True)
        excel.Quit()

    except Exception as e:
        print(f"Error checking background query status: {e}")

# Usage
check_if_query_is_running("C:\\Users\\kmitsios\\Desktop\\Enrollments Data.xlsx")
