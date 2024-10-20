# Program to read a file and display the contents in a table format
# Author: Frode Rørvik
# Date: 2024-10-13

import pandas as pd
import os
from datetime import datetime
import openpyxl


def check_file_in_use(file_path):
    if not os.path.exists(file_path):
        return False
    try:
        # Try to open the file in exclusive mode
        with open(file_path, 'r+'):
            pass
    except OSError:
        print(f"\033[91mError: The file {file_path} is in use by another application. Please close the application\033[0m")
        return True
    return False


def read_file_to_table():
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    default_filename = os.path.join(downloads_dir, "SALARY.SI")
    filename = input(f"Enter the filename (default is {default_filename}): ") or default_filename
    output_filename = os.path.join(os.path.dirname(filename), "SALARY.xlsx")

    if check_file_in_use(output_filename) or check_file_in_use(filename):
        return
    
    try:
        with open(filename, 'r', encoding='latin1') as file:
            lines = file.readlines()
        
        # Extract RunDate from the appropriate line
        run_date_line = lines[3].strip()
        run_date_str = run_date_line.split()[1]  # Extract the date string after #GEN
        RunDate = datetime.strptime(run_date_str, "%Y%m%d").date()  # Convert to date datatype
        
        # Remove paragraph marks, replace {} with {"", "", "", ""}, and store the remaining lines in an array
        data_lines = [line.strip().replace('{}', '{"" "" "" ""}').replace('{', '').replace('}', '') for line in lines[15:] if line.strip() != '']
        
        # Transform data lines to array
        data_array = transform_to_array(data_lines)
        
        # Convert data array to DataFrame
        df = pd.DataFrame(data_array)

         # Remove the first column
        df = df.drop(df.columns[[0, 2, 3, 4, 6, 7, 8 ]], axis=1)
        df = df.drop(df.columns[[5, 6, 7 ]], axis=1)
        # Convert numeric columns to appropriate data types
        # df = df.apply(pd.to_numeric, errors='ignore')
        df.iloc[:, 0] = pd.to_numeric(df.iloc[:, 0], errors='coerce')
        df.iloc[:, 2] = pd.to_numeric(df.iloc[:, 2], errors='coerce')
        df.iloc[:, 3] = pd.to_numeric(df.iloc[:, 3], errors='coerce')
        df.iloc[:, 4] = pd.to_datetime(df.iloc[:, 4], format='%Y%m%d', errors='coerce')
        
        # Write the output to an Excel file
        
        df.to_excel(output_filename, index=False, header=False)
        
        print(f"RunDate: {RunDate}")
        print(f"Data has been written to {output_filename}")
        
    except FileNotFoundError:
        print(f"File {filename} not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

def transform_to_array(data_lines):
    data_array = []
    for line in data_lines:
        # Split the line by spaces, but keep quoted strings together
        parts = []
        current_part = []
        in_quotes = False
        previous_char_was_space = False
        for char in line:
            if char == '"':
                if in_quotes:
                    # End of quoted string
                    parts.append(''.join(current_part))
                    current_part = []
                else:
                    # Start of quoted string
                    if current_part:
                        parts.append(''.join(current_part))
                        current_part = []
                    parts.append('')  # Add empty string for blank value
                in_quotes = not in_quotes
                previous_char_was_space = False
            elif char == ' ' and not in_quotes:
                if not previous_char_was_space:
                    if current_part:
                        parts.append(''.join(current_part))
                        current_part = []
                    previous_char_was_space = True
            else:
                current_part.append(char)
                previous_char_was_space = False
        if current_part:
            parts.append(''.join(current_part))
        data_array.append(parts)
    return data_array

if __name__ == "__main__":
    read_file_to_table()