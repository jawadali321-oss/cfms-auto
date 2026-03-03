#!/usr/bin/env python3
import os
from datetime import datetime

# Path to the local files in the Documents folder
WATCH_FILES = {
    "framing_errors.txt": "~/Documents/framing_errors.txt",
    "invalidcases.txt": "~/Documents/invalidcases.txt",
    "no_final_order_cases.txt": "~/Documents/no_final_order_cases.txt",
}

# Function to expand the ~ to the full home directory path
def expand(path):
    return os.path.expanduser(path)

# Function to append a message to a file
def append_to_file(filepath, message):
    try:
        with open(filepath, 'a') as file:
            file.write(message + "\n")
            print(f"✅ Data written to {filepath}")
    except Exception as e:
        print(f"❌ Error writing to {filepath}: {e}")

# Example function to send a test message to the files
def send_test_data():
    message = f"Test data sent at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    
    # Sending the message to each file in the WATCH_FILES
    for file_name, file_path in WATCH_FILES.items():
        full_path = expand(file_path)
        append_to_file(full_path, message)

if __name__ == "__main__":
    send_test_data()
