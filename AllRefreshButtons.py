import tkinter as tk

def button_click(button_name):
    print(f"Button clicked: {button_name}")

root = tk.Tk()
root.title("Button Box")

# Create label frames for each group of buttons
frame_12 = tk.LabelFrame(root, text="1 & 2 Buttons")
frame_3 = tk.LabelFrame(root, text="3 Button")
frame_456 = tk.LabelFrame(root, text="4 & 5 & 6 Buttons")
frame_789 = tk.LabelFrame(root, text="7 & 8 & 9 Buttons")

# Pack the label frames
frame_12.pack(padx=10, pady=5, fill=tk.X)
frame_3.pack(padx=10, pady=5, fill=tk.X)
frame_456.pack(padx=10, pady=5, fill=tk.X)
frame_789.pack(padx=10, pady=5, fill=tk.X)

# Create buttons and add them to their respective frames
button_names = [
    ("Individuals", frame_12),
    ("Contacts", frame_12),
    ("HELIOS DB", frame_3),
    ("Outreach v2", frame_456),
    ("Partners Outreach", frame_456),
    ("Accommodation Outreach", frame_456),
    ("Accommodation Services csv", frame_789),
    ("Matchmaking csv", frame_789),
    ("MD csv", frame_789),
]

for button_name, frame in button_names:
    button = tk.Button(frame, text=button_name, command=lambda name=button_name: button_click(name))
    button.pack(side=tk.LEFT, padx=5, pady=5)

root.mainloop()
