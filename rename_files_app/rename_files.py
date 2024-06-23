import zipfile
import pandas as pd
import os
import re

# Step 1: Extract files from the zip archive
zip_file_path = 'test.zip'
extract_path = 'extracted_files'

with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
    zip_ref.extractall(extract_path)

# List extracted files
extracted_files = os.listdir(extract_path)
print("Extracted files:", extracted_files)

# Step 2: Read the Excel file (bank statement)
excel_files = [f for f in extracted_files if f.endswith('.xlsx')]
bank_statement_df = pd.read_excel(os.path.join(extract_path, excel_files[0]))
print("Bank statement:")
print(bank_statement_df.head())

# Step 3: Read PDF and Image files
pdf_files = [f for f in extracted_files if f.endswith('.pdf')]
jpg_files = [f for f in extracted_files if f.endswith('.jpg')]

def generate_new_file_name(posting_date, description, amount, file_extension):
    merchant = description.replace(" ", "_")
    amount_str = f"{amount:.2f}"
    new_file_name = f"{posting_date}_{merchant}_{amount_str}.{file_extension}"
    return new_file_name

def extract_date_amount_from_filename(filename):
    pattern = r"(\d{4}-\d{2}-\d{2})_(\d+\.\d{2})"
    match = re.match(pattern, filename)
    if match:
        date, amount = match.groups()
        return date, float(amount)
    return None, None

# Function to match files and rename them
def rename_files(file_list, file_extension):
    for file in file_list:
        base_name, _ = os.path.splitext(file)
        date, amount = extract_date_amount_from_filename(base_name)
        
        if date and amount:
            matching_entries = bank_statement_df[
                (bank_statement_df['Posting Date'] == date) &
                (bank_statement_df['Amount'] == amount)
            ]
            
            if not matching_entries.empty:
                matching_entry = matching_entries.iloc[0]
                new_file_name = generate_new_file_name(
                    matching_entry['Posting Date'],
                    matching_entry['Description'],
                    matching_entry['Amount'],
                    file_extension
                )
                old_file_path = os.path.join(extract_path, file)
                new_file_path = os.path.join(extract_path, new_file_name)
                os.rename(old_file_path, new_file_path)
                print(f"Renamed {file} to {new_file_name}")
            else:
                print(f"No match found for {file}")

# Rename PDF files
rename_files(pdf_files, 'pdf')

# Rename JPG files
rename_files(jpg_files, 'jpg')
