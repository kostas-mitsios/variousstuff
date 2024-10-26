import tkinter as tk
from tkinter import ttk

def open_donations_menu():
    donations_menu_window = tk.Toplevel(root)
    donations_menu_window.title("Donations Menu")
    donations_menu_window.configure(bg="blue")  # Set the background color to blue

    input_frame = ttk.Frame(donations_menu_window, padding=20)
    input_frame.grid(row=0, column=0, sticky="nsew")

    view_donations_button = ttk.Button(input_frame, text="View Donations", command=on_view_donations)
    view_donations_button.grid(row=1, column=0, columnspan=1, pady=10)

    add_donation_button = ttk.Button(input_frame, text="Add Donations", command=on_add_donations)
    add_donation_button.grid(row=1, column=2, columnspan=1, pady=10)

    return_button = ttk.Button(input_frame, text="Return", command=donations_menu_window.destroy)
    return_button.grid(row=2, column=1, pady=10)

    input_frame.columnconfigure((1, 2), weight=1)

def on_view_donations():
    print("View Donations")

def on_add_donations():
    print("Add Donations")

root = tk.Tk()
root.title("Main Window")
root.configure(bg="blue")  # Set the background color of the main window to blue

donations_button = ttk.Button(root, text="Open Donations Menu", command=open_donations_menu)
donations_button.pack(pady=10)

root.mainloop()
