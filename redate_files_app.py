import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import shutil
import re
from datetime import datetime


class FileRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Renamer App")
        
        self.excel_file_path = ""
        self.input_folder_path = ""
        self.matched_folder_path = ""
        self.unmatched_folder_path = ""

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Excel File:").grid(row=0, column=0, padx=10, pady=10)
        self.excel_file_entry = tk.Entry(self.root, width=50)
        self.excel_file_entry.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.browse_excel_file).grid(row=0, column=2, padx=10, pady=10)
        
        tk.Label(self.root, text="Input Folder:").grid(row=1, column=0, padx=10, pady=10)
        self.input_folder_entry = tk.Entry(self.root, width=50)
        self.input_folder_entry.grid(row=1, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.browse_input_folder).grid(row=1, column=2, padx=10, pady=10)
        
        tk.Button(self.root, text="Start", command=self.start_processing).grid(row=2, column=1, pady=20)

    def browse_excel_file(self):
        self.excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.excel_file_entry.delete(0, tk.END)
        self.excel_file_entry.insert(0, self.excel_file_path)

    def browse_input_folder(self):
        self.input_folder_path = filedialog.askdirectory()
        self.input_folder_entry.delete(0, tk.END)
        self.input_folder_entry.insert(0, self.input_folder_path)

    def log_message(self, message):
        log_file_path = os.path.join(self.unmatched_folder_path, "file_processing_log.txt")
        with open(log_file_path, "a") as log_file:
            log_file.write(f"{datetime.now()} - {message}\n")

    def start_processing(self):
        try:
            self.matched_folder_path = os.path.join(self.input_folder_path, "matched_files")
            self.unmatched_folder_path = os.path.join(self.input_folder_path, "unmatched_files")
            os.makedirs(self.matched_folder_path, exist_ok=True)
            os.makedirs(self.unmatched_folder_path, exist_ok=True)

            df = pd.read_excel(self.excel_file_path)
            df['Amount'] = df['Amount'].apply(lambda x: round(x, 2))

            # extract amount from file name
            def extract_amount(file_name):
                try:
                    # formats: underscore and dash
                    amount_part = re.search(r'[\$_-](\d+[.,]\d{2})', file_name)
                    if amount_part:
                        amount_str = amount_part.group(1).replace(',', '.')
                        return float(amount_str)
                    return None
                except ValueError:
                    return None

            for file in os.listdir(self.input_folder_path):
                if file.endswith(('.pdf', '.jpg', '.jpeg', '.png')):
                    file_path = os.path.join(self.input_folder_path, file)
                    amount = extract_amount(file)
                    
                    if amount is not None:
                        matched_rows = df[df['Amount'] == amount]
                        
                        if not matched_rows.empty:
                            posting_date = pd.to_datetime(matched_rows.iloc[0]['Posting Date']).strftime('%m%d%Y')
                            
                            file_extension = os.path.splitext(file)[1]
                            
                            remaining_part_match = re.search(r'[_-](.*)', file)
                            remaining_part = remaining_part_match.group(1) if remaining_part_match else file
                            
                            new_file_name = f"{posting_date}_{remaining_part}"
                            new_file_path = os.path.join(self.matched_folder_path, new_file_name)
                            shutil.copy(file_path, new_file_path)
                            self.log_message(f"Matched and renamed: {file} -> {new_file_name}")
                        else:
                            shutil.copy(file_path, os.path.join(self.unmatched_folder_path, file))
                            self.log_message(f"No match found for: {file}")
                    else:
                        shutil.copy(file_path, os.path.join(self.unmatched_folder_path, file))
                        self.log_message(f"Invalid amount in file name: {file}")

            messagebox.showinfo("Success", "Files processed successfully!")
            self.log_message("Processing completed successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.log_message(f"Error: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileRenamerApp(root)
    root.mainloop()
