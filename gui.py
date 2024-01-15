import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
from tkinter import filedialog, scrolledtext
from functions import extract_reference_ids, highlight_matches_in_excel
import sys

class StdoutRedirector(object):
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.config(state=tk.NORMAL)  # Enable the text widget
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)  # Auto-scroll to the bottom
        self.text_widget.config(state=tk.DISABLED)  # Disable the text widget

    def flush(self):
        pass  # This might be needed if using Python 2.x


def select_folder():
    folder_path = filedialog.askdirectory()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, folder_path)

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

def start_processing():
    output_text.delete(1.0, tk.END)  # Clear the text field before new output
    folder_path = folder_entry.get()
    file_path = file_entry.get()
    if folder_path and file_path:
        reference_ids = extract_reference_ids(folder_path)
        highlight_matches_in_excel(file_path, reference_ids)
        print("Done!")
    else:
        print("Please select both a folder and a file.")
        
def on_start_button_pressed():
    print("Processing...")
    root.update_idletasks()  # Update the GUI to reflect the print statement
    start_processing()

# Set up the main application window
root = ThemedTk(theme="adapta")
root.title("Excel Processing App")

# Folder selection
tk.Label(root, text="Select folder with all doubletten analysis files:").pack()
folder_entry = ttk.Entry(root, width=50)
folder_entry.pack()
ttk.Button(root, text="Browse", command=select_folder).pack()

# File selection
tk.Label(root, text="Select Excel file to compare against:").pack()
file_entry = ttk.Entry(root, width=50)
file_entry.pack()
ttk.Button(root, text="Browse", command=select_file).pack()

# Start button
ttk.Button(root, text="Start", command=on_start_button_pressed).pack()

# Text widget for displaying print statements
output_text = scrolledtext.ScrolledText(root, height=10)
output_text.pack()
output_text.config(state=tk.DISABLED)  # Start with the text widget disabled


# Redirect standard output
sys.stdout = StdoutRedirector(output_text)

# Run the application
root.mainloop()
