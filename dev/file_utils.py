import os
from datetime import datetime

def log_message(unmatched_folder_path, message):
    log_file_path = os.path.join(unmatched_folder_path, "file_processing_log.txt")
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{datetime.now()} - {message}\n")
