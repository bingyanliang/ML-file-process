import os
import pandas as pd
import re
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox

def extract_zip(zip_path, extract_to):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)

def load_excel_data(file_path):
    excel_data = pd.read_excel(file_path)
    return excel_data[['Posting Date', 'Amount', 'Description']]

def rename_files(directory, excel_data):
    for root, _, files in os.walk(directory):
        for file in files:
            match = re.match(r'(\d{8})_\$(\d+).(pdf|jpg)', file)
            if match:
                date, amount, ext = match.groups()
                amount = amount.replace(',', '')  # Remove commas for consistency
                
                entry = excel_data[(excel_data['Posting Date'] == date) & 
                                   (excel_data['Amount'] == float(amount))]
                if not entry.empty:
                    merchant = entry['Description'].values[0]
                    merchant_clean = re.sub(r'[^\w\s]', '', merchant).replace(' ', '_')
                    new_name = f"{date}_{merchant_clean}_${amount}.{ext}"
                    os.rename(os.path.join(root, file), os.path.join(root, new_name))
                    print(f"Renamed {file} to {new_name}")

def zip_directory(directory, zip_path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(directory):
            for file in files:
                zipf.write(os.path.join(root, file), 
                           os.path.relpath(os.path.join(root, file), 
                           os.path.join(directory, '..')))

def select_zip_file():
    file_path = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")])
    zip_path_var.set(file_path)

def execute_renaming():
    zip_file_path = zip_path_var.get()
    
    if not zip_file_path:
        messagebox.showerror("Error", "Please select a ZIP file")
        return
    
    extract_to = os.path.join(os.path.dirname(zip_file_path), 'extracted_files')
    os.makedirs(extract_to, exist_ok=True)
    
    try:
        # Extract ZIP file
        extract_zip(zip_file_path, extract_to)
        
        # Locate Excel file
        excel_file_path = None
        for root, _, files in os.walk(extract_to):
            for file in files:
                if file.endswith('.xlsx'):
                    excel_file_path = os.path.join(root, file)
                    break
        
        if not excel_file_path:
            messagebox.showerror("Error", "No Excel file found in the ZIP")
            return
        
        # Load Excel data
        excel_data = load_excel_data(excel_file_path)
        
        # Rename files
        rename_files(extract_to, excel_data)
        
        # Re-zip the directory
        zip_output_path = os.path.join(os.path.dirname(zip_file_path), 'renamed_files.zip')
        zip_directory(extract_to, zip_output_path)
        
        messagebox.showinfo("Success", f"Files renamed and zipped to {zip_output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Initialize the Tkinter window
root = tk.Tk()
root.title("Invoice Renaming Tool")

# Variable to store file path
zip_path_var = tk.StringVar()

# Create UI elements
tk.Label(root, text="Select ZIP File:").grid(row=0, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=zip_path_var, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_zip_file).grid(row=0, column=2, padx=10, pady=10)

tk.Button(root, text="Rename and Zip Files", command=execute_renaming).grid(row=1, columnspan=3, pady=20)

# Run the Tkinter event loop
root.mainloop()
