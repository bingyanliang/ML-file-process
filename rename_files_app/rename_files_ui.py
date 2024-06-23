import os
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox

def load_excel_data(file_path):
    excel_data = pd.read_excel(file_path)
    # Ensure Posting Date is in MM/DD/YYYY format
    excel_data['Posting Date'] = pd.to_datetime(excel_data['Posting Date']).dt.strftime('%m/%d/%Y')
    # Remove leading/trailing spaces from Description
    excel_data['Description'] = excel_data['Description'].str.strip()
    # Extract integer part of Amount
    excel_data['Integer Amount'] = excel_data['Amount'].astype(int)
    print("Loaded Excel Data:")
    print(excel_data.head())
    return excel_data[['Posting Date', 'Integer Amount', 'Description']]

def rename_files(directory, excel_data):
    files = os.listdir(directory)
    print(f"Files in directory {directory}: {files}")
    
    for file in files:
        # Match the file pattern MMDDYYYY_$amount.ext
        match = re.match(r'(\d{2})(\d{2})(\d{4})_\$(\d+)\.(pdf|jpg)', file)
        if match:
            month, day, year, amount_str, ext = match.groups()
            date = f"{month}/{day}/{year}"
            amount = int(amount_str)
            print(f"Processing file: {file}, extracted date: {date}, amount: {amount}")
            
            entry = excel_data[(excel_data['Posting Date'] == date) & 
                               (excel_data['Integer Amount'] == amount)]
            
            print(f"Matching entries for date {date} and amount {amount}:")
            print(entry)
            
            if not entry.empty:
                merchant = entry['Description'].values[0]
                # Clean merchant name
                merchant_clean = re.sub(r'[^\w\s]', '', merchant).replace(' ', '_')
                new_name = f"{month}-{day}-{year}_{merchant_clean}_{amount}.{ext}"
                old_file_path = os.path.join(directory, file)
                new_file_path = os.path.join(directory, new_name)
                os.rename(old_file_path, new_file_path)
                print(f"Renamed {file} to {new_name}")
            else:
                print(f"No matching entry found for file: {file}")
        else:
            print(f"File pattern did not match for file: {file}")

def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    excel_path_var.set(file_path)

def select_directory():
    directory = filedialog.askdirectory()
    directory_path_var.set(directory)

def execute_renaming():
    excel_file_path = excel_path_var.get()
    files_directory = directory_path_var.get()
    
    if not excel_file_path or not files_directory:
        messagebox.showerror("Error", "Please select both an Excel file and a directory")
        return
    
    try:
        excel_data = load_excel_data(excel_file_path)
        rename_files(files_directory, excel_data)
        messagebox.showinfo("Success", "Files renamed successfully")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Initialize the Tkinter window
root = tk.Tk()
root.title("File Renaming Tool")

# Variables to store file paths
excel_path_var = tk.StringVar()
directory_path_var = tk.StringVar()

# Create UI elements
tk.Label(root, text="Select Excel File:").grid(row=0, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=excel_path_var, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_excel_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Select Directory:").grid(row=1, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=directory_path_var, width=50).grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_directory).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Rename Files", command=execute_renaming).grid(row=2, columnspan=3, pady=20)

# Run the Tkinter event loop
root.mainloop()

