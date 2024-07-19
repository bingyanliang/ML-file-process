import os
import re
import shutil
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox
from excel_utils import read_excel, write_excel, update_category_col, apply_row_colors

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
        from tkinter import Label, Entry, Button

        Label(self.root, text="Excel File:").grid(row=0, column=0, padx=10, pady=10)
        self.excel_file_entry = Entry(self.root, width=50)
        self.excel_file_entry.grid(row=0, column=1, padx=10, pady=10)
        Button(self.root, text="Browse", command=self.browse_excel_file).grid(row=0, column=2, padx=10, pady=10)
        
        Label(self.root, text="Input Folder:").grid(row=1, column=0, padx=10, pady=10)
        self.input_folder_entry = Entry(self.root, width=50)
        self.input_folder_entry.grid(row=1, column=1, padx=10, pady=10)
        Button(self.root, text="Browse", command=self.browse_input_folder).grid(row=1, column=2, padx=10, pady=10)
        
        Button(self.root, text="Start", command=self.start_processing).grid(row=2, column=1, pady=20)
        Button(self.root, text="Update Category", command=self.update_category).grid(row=3, column=1, pady=10)

    def browse_excel_file(self):
        self.excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.excel_file_entry.delete(0, tk.END)
        self.excel_file_entry.insert(0, self.excel_file_path)

    def browse_input_folder(self):
        self.input_folder_path = filedialog.askdirectory()
        self.input_folder_entry.delete(0, tk.END)
        self.input_folder_entry.insert(0, self.input_folder_path)

    def start_processing(self):
        try:
            if not self.excel_file_path or not self.input_folder_path:
                raise ValueError("Please select both an Excel file and an input folder.")
            
            self.matched_folder_path = os.path.join(self.input_folder_path, "matched_files")
            self.unmatched_folder_path = os.path.join(self.input_folder_path, "unmatched_files")
            os.makedirs(self.matched_folder_path, exist_ok=True)
            os.makedirs(self.unmatched_folder_path, exist_ok=True)

            headers, data, wb, ws = read_excel(self.excel_file_path)

            # Add a new column for the match status
            if 'Matched' not in headers:
                headers = list(headers)  # Convert headers tuple to list
                headers.append('Matched')
                ws.cell(row=1, column=len(headers), value='Matched')

            # Create a list of dictionaries for easier access
            df = [
                {headers[i]: row[i] if i < len(row) else '' for i in range(len(headers))}
                for row in data
            ]

            for file in os.listdir(self.input_folder_path):
                if file.endswith(('.pdf', '.jpg', '.jpeg', '.png')):
                    file_path = os.path.join(self.input_folder_path, file)
                    self.process_file(file_path, file, df, headers)

            # Write the updated data back to the Excel file 
            write_excel(self.excel_file_path, headers, df, wb, ws)

            messagebox.showinfo("Success", "Files processed successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def process_file(self, file_path, file, df, headers):
        # Split the file name into parts
        file_name, file_extension = os.path.splitext(file)
        parts = re.split(r'[_-]', file_name)
        
        if len(parts) != 3:
            shutil.copy(file_path, os.path.join(self.unmatched_folder_path, file))
            return

        original_date, middle_part, amount_part = parts

        # Remove the comma from the amount part for matching
        amount_str = amount_part.replace(',', '').replace('$', '')
        try:
            amount = float(amount_str)
        except ValueError:
            shutil.copy(file_path, os.path.join(self.unmatched_folder_path, file))
            return

        matched_rows = [row for row in df if row['Amount'] == amount]
        
        if matched_rows:
            posting_date = matched_rows[0]['Posting Date']
            if isinstance(posting_date, datetime):
                posting_date_str = posting_date.strftime('%m%d%Y')
            else:
                posting_date_str = datetime.strptime(posting_date, '%Y/%m/%d').strftime('%m%d%Y')

            formatted_amount = f"${amount:,.2f}"
            new_file_name = f"{posting_date_str}_{middle_part}_{formatted_amount}{file_extension}"
            
            new_file_path = os.path.join(self.matched_folder_path, new_file_name)
            shutil.copy(file_path, new_file_path)

            # Update the matched rows with 'Y'
            for row in df:
                if row['Amount'] == amount:
                    row['Matched'] = 'Y'
        else:
            shutil.copy(file_path, os.path.join(self.unmatched_folder_path, file))

    def update_category(self):
        try: 
            if not self.excel_file_path:
                raise ValueError("Please select an excel file")
            
            headers, data, wb, ws = read_excel(self.excel_file_path)
            headers = list(headers)
            update_category_col(headers, data, wb, ws)
            apply_row_colors(headers, data, wb, ws)
            wb.save(self.excel_file_path)
            messagebox.showinfo("Success", "Category updated successfully!")

        except Exception as e:
            messagebox.showerror('Error', str(e))
