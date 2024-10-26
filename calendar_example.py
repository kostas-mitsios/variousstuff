import tkinter as tk
from tkcalendar import Calendar
import gspread
from google.oauth2 import service_account

labels = ["name", "date_started", "date_stopped", "usual_shift_day", "usual_shift_time", "tourist"]

def show_info():
    selected_date = cal.get_date()
    info = get_data_for_date(selected_date)
    
    if info:
        display_info(selected_date, info)
    else:
        display_info(selected_date, "No data found for this date.")

def get_data_for_date(selected_date):
    # Authenticate with Google Sheets API using your credentials
    credentials_file = 'credentials.json'
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    credentials = service_account.Credentials.from_service_account_file(credentials_file, scopes=scope)
    client = gspread.Client(auth=credentials)
    
    # Open the Google Sheet by its ID
    sheet_id = "1sm_ePXHiTxty5gHfgW_rWGhzdnml1G6sVfuwKfKWeMM"
    worksheet = client.open_by_key(sheet_id).worksheet("Volunteers")
    
    # Fetch all data from column B
    column_b_data = worksheet.col_values(2)
    
    # Collect all rows with matching dates
    matching_rows = []

    for i, date in enumerate(column_b_data):
        if date == selected_date:
            row_with_date = worksheet.row_values(i + 1)  # Adjust for 0-based indexing
            matching_rows.append(row_with_date)

    formatted_data_list = []
    

    for row in matching_rows:
        formatted_data = '\n'.join([f"{label}: {value}" for label, value in zip(labels, row)])
        formatted_data_list.append(formatted_data)

    # Join the formatted data with newlines to get a single string
    formatted_data_string = '\n\n'.join(formatted_data_list)

    return formatted_data_string

def display_info(selected_date, info):
    info_window = tk.Toplevel(app)
    info_window.title("Information")
    
    info_text = tk.Text(info_window, wrap=tk.WORD)
    info_text.pack(padx=10, pady=10)

    info_text.insert(tk.END, f"Data for {selected_date}:\n\n{''.join(info)}")

    # Set the color for the words in the labels
    for label in labels:
        start = info_text.search(label, "1.0", stop=tk.END)
        end = f"{start}+{len(label)}c"
        info_text.tag_add(label, start, end)
        info_text.tag_config(label, foreground="blue")  # Change the color here

app = tk.Tk()
app.title("Calendar Example")

# Create a Calendar widget
cal = Calendar(app, date_pattern="dd/mm/y")
cal.pack(padx=10, pady=10)

# Button to show information for the selected date
show_info_button = tk.Button(app, text="Show Info", command=show_info)
show_info_button.pack(pady=10)

app.mainloop()
