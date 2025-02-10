import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from modules.dbp import dbp_statement_txt_excel

methods = [
    ("DBP STATEMENT TXT TO XLSX", dbp_statement_txt_excel)
]

# Function to browse for a file
def browse_file():
    filename = filedialog.askopenfilename(title="Select a file")
    file_entry.delete(0, tk.END)  # Clear current text
    file_entry.insert(0, filename)  # Insert selected filename

# Function to handle method selection
def on_method_change(event):
    selected_method = method_combobox.get()
    # print(f"Selected Method: {selected_method}")

# Function to save output (just as an example action)
def process_data():
    input_file = file_entry.get()
    selected_method = method_combobox.get()
    output_name = output_entry.get()
    
    if not input_file or not selected_method or not output_name:
        messagebox.showerror('Error', 'Please fill all fields')
        return
    
    # print(f"Processing file: {input_file}")
    # print(f"Using method: {selected_method}")
    # print(f"Output file will be saved as: {output_name}")
    
    # Add your processing logic here
    # For example, creating the output file:
    
    for method in methods:
        if(selected_method == method[0]):
            process = method[1](input_file, output_name)
            if process[0]: messagebox.showinfo("Success", process[1])
            else: messagebox.showerror("Error", process[1])
            break

# Create the main window
root = tk.Tk()
root.title("PGH File Processor")
root.iconbitmap("assets/UP_PGH_logo.ico")

# Create a ttk.Style object for modern styling
style = ttk.Style()

# Set the window size and dark theme background color
root.geometry("500x150")
root.configure(bg="#2e2e2e")  # Dark background for the main window

# Define dark theme colors
bg_color = "#2e2e2e"
fg_color = "#ffffff"
button_bg = "#444444"
button_fg = "#ffffff"
entry_bg = "#444444"
entry_fg = "#ffffff"
combobox_bg = "#444444"
combobox_fg = "#ffffff"

# Method selection label and combobox
method_label = tk.Label(root, text="Select Method:", bg=bg_color, fg=fg_color, font=("Helvetica", 10))
method_label.grid(row = 0, column = 0, sticky="w", padx=10, pady=5)

method_combobox = ttk.Combobox(root, values=[i[0] for i in methods], state="readonly", 
                                background=entry_bg, foreground=entry_fg, width=32, font=("Helvetica", 10), style="TCombobox")

method_combobox.grid(row = 0, column = 1, sticky="w", padx=10, pady=5)
method_combobox.bind("<<ComboboxSelected>>", on_method_change)

# File selection label and entry field
file_label = tk.Label(root, text="Select Input File:", bg=bg_color, fg=fg_color, font=("Helvetica", 10))
file_label.grid(row = 1, column = 0, sticky="w", padx=10, pady=5)

file_entry = ttk.Entry(root, width=40, style="TEntry")
file_entry.grid(row = 1, column = 1, sticky="w", padx=10, pady=5)

browse_button = ttk.Button(root, text="Browse", command=browse_file, style="TButton")
browse_button.grid(row = 1, column = 3, sticky="w", padx=10, pady=5)

# Output name field
output_label = tk.Label(root, text="Output File Name:", bg=bg_color, fg=fg_color, font=("Helvetica", 10))
output_label.grid(row = 2, column = 0, sticky="w", padx=10, pady=5)

output_entry = ttk.Entry(root, width=40)
output_entry.grid(row = 2, column = 1, sticky="w", padx=10, pady=5)

# Process Button
process_button = ttk.Button(root, text="Process", command=process_data, style="TButton")
process_button.grid(row = 3, column = 0, sticky="w", padx=10, pady=5)

# Start the Tkinter event loop
root.mainloop()
