import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import os
import shutil
import re
from datetime import datetime

class FileRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Renamer App")
        
        self.csv_file_path = ""
        self.input_folder_path = ""
        self.matched_folder_path = ""
        self.unmatched_folder_path = ""

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="CSV File:").grid(row=0, column=0, padx=10, pady=10)
        self.csv_file_entry = tk.Entry(self.root, width=50)
        self.csv_file_entry.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.browse_csv_file).grid(row=0, column=2, padx=10, pady=10)
        
        tk.Label(self.root, text="Input Folder:").grid(row=1, column=0, padx=10, pady=10)
        self.input_folder_entry = tk.Entry(self.root, width=50)
        self.input_folder_entry.grid(row=1, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.browse_input_folder).grid(row=1, column=2, padx=10, pady=10)
        
        tk.Button(self.root, text="Start", command=self.start_processing).grid(row=2, column=1, pady=20)

    def browse_csv_file(self):
        self.csv_file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.csv_file_entry.delete(0, tk.END)
        self.csv_file_entry.insert(0, self.csv_file_path)

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
            if not self.csv_file_path or not self.input_folder_path:
                raise ValueError("Please select both a CSV file and an input folder.")
            
            self.matched_folder_path = os.path.join(self.input_folder_path, "matched_files")
            self.unmatched_folder_path = os.path.join(self.input_folder_path, "unmatched_files")
            os.makedirs(self.matched_folder_path, exist_ok=True)
            os.makedirs(self.unmatched_folder_path, exist_ok=True)

            records = []
            with open(self.csv_file_path, mode='r', newline='') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    row['Amount'] = round(float(row['Amount']), 2)
                    records.append(row)

            def process_file(file_path, file):
                # convert file name into 3 parts divided by 2 underscores
                file_name, file_extension = os.path.splitext(file)
                parts = re.split(r'[_-]', file_name)
                
                if len(parts) != 3:
                    shutil.copy(file_path, os.path.join(self.unmatched_folder_path, file))
                    self.log_message(f"Invalid file name format: {file}")
                    return

                original_date, middle_part, amount_part = parts

                # Remove the comma from amount
                amount_str = amount_part.replace(',', '').replace('$', '')
                try:
                    amount = float(amount_str)
                except ValueError:
                    shutil.copy(file_path, os.path.join(self.unmatched_folder_path, file))
                    self.log_message(f"Invalid amount in file name: {file}")
                    return

                matched_row = next((row for row in records if row['Amount'] == amount), None)
                
                if matched_row:
                    posting_date = datetime.strptime(matched_row['Posting Date'], '%m/%d/%Y').strftime('%m%d%Y')
                    formatted_amount = f"${amount:,.2f}"
                    new_file_name = f"{posting_date}_{middle_part}_{formatted_amount}{file_extension}"
                    
                    new_file_path = os.path.join(self.matched_folder_path, new_file_name)
                    shutil.copy(file_path, new_file_path)
                    self.log_message(f"Matched and renamed: {file} -> {new_file_name}")
                else:
                    shutil.copy(file_path, os.path.join(self.unmatched_folder_path, file))
                    self.log_message(f"No match found for: {file}")

            for file in os.listdir(self.input_folder_path):
                if file.endswith(('.pdf', '.jpg', '.jpeg', '.png')):
                    file_path = os.path.join(self.input_folder_path, file)
                    process_file(file_path, file)

            messagebox.showinfo("Success", "Files processed successfully!")
            self.log_message("Processing completed successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.log_message(f"Error: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileRenamerApp(root)
    root.mainloop()
