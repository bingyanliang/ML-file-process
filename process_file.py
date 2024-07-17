import re
import shutil
import datetime
import os

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

    else:
        shutil.copy(file_path, os.path.join(unmatched_folder_path, file))