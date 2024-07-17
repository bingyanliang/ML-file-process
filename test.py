import pylightxl as xl
import os
import shutil
import re
from datetime import datetime


excel_file_path = "test\CreditCardAsof05312024.xlsx"
input_folder_path = "test"

try:

    matched_folder_path = os.path.join(input_folder_path, "matched_files")
    unmatched_folder_path = os.path.join(input_folder_path, "unmatched_files")
    os.makedirs(matched_folder_path, exist_ok=True)
    os.makedirs(unmatched_folder_path, exist_ok=True)

    # Read the Excel file using pylightxl
    db = xl.readxl(fn=excel_file_path)
    ws = db.ws(db.ws_names[0])
    data = list(ws.rows)

    # Extract header and data
    headers = data[0]
    data = data[1:]

    # add a column for matched files
    if 'Matched' not in headers:
        headers.append('Matched')

    # Create a list of dictionaries for easier access
    df = [{headers[i]: row[i] for i in range(len(headers))} for row in data]

    # Process each file in the input folder
    def process_file(file_path, file):
        # Split the file name into parts
        file_name, file_extension = os.path.splitext(file)
        parts = re.split(r'[_-]', file_name)
                
        if len(parts) != 3:
            shutil.copy(file_path, os.path.join(unmatched_folder_path, file))
            return

        original_date, middle_part, amount_part = parts
        # Remove the comma from the amount part for matching
        amount_str = amount_part.replace(',', '').replace('$', '')
        try:
            amount = float(amount_str)
        except ValueError:
            shutil.copy(file_path, os.path.join(unmatched_folder_path, file))
            return

        matched_rows = [row for row in df if row['Amount'] == amount]
                
        if matched_rows:
            posting_date = datetime.strptime(matched_rows[0]['Posting Date'], '%Y/%m/%d').strftime('%m%d%Y')
            formatted_amount = f"${amount:,.2f}"
            new_file_name = f"{posting_date}_{middle_part}_{formatted_amount}{file_extension}"
                    
            new_file_path = os.path.join(matched_folder_path, new_file_name)
            shutil.copy(file_path, new_file_path)

            # update matched marks
            for row in df:
                if row['Amount'] == amount:
                    row['Matched'] = 'Y'

            for file in os.listdir(input_folder_path):
                if file.endswith(('.pdf', '.jpg', '.jpeg', '.png')):
                    file_path = os.path.join(input_folder_path, file)
                    process_file(file_path, file)

            # Write the updated data back to the Excel file
            new_data = [headers] + [[row[header] for header in headers] for row in df]
            db.ws(ws.name).update_index(row=1, col=1, val=new_data)

            xl.writexl(db=db, fn=excel_file_path)

except Exception as e:
    None
