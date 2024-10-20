import pandas as pd
import os
import warnings
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime


def check_file_in_use(file_path):
    try:
        # Try to open the file in exclusive mode
        with open(file_path, 'r+'):
            pass
    except OSError:
        print(f"\033[91mError: The file {file_path} is in use by another application. Please close the application\033[0m")
        return True
    return False

def read_and_splice_excel_files():
    # Suppress warnings
    warnings.simplefilter(action='ignore', category=UserWarning)
    
    # Define the file paths
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    detaljert_file = os.path.join(downloads_dir, "Transaksjoner, detaljert.xlsx")
    konto_file = os.path.join(downloads_dir, "Transaksjoner per konto og kostnadsbærer.xlsx")

    # Check if any of the files are in use
    if check_file_in_use(detaljert_file) or check_file_in_use(konto_file):
        return
    
    # Get today's date and format it
    today_date = datetime.today().strftime("%d.%m.%Y")
    output_file = os.path.join(downloads_dir, f"splice {today_date}.xlsx")

    if check_file_in_use(output_file):
        return
    
    # Read the Excel files, skipping the first two rows
    detaljert_df = pd.read_excel(detaljert_file, header=0, skiprows=[0, 1])
    konto_df = pd.read_excel(konto_file, header=0, skiprows=[0, 1])
    
    # Select the required columns from konto.xlsx by their actual names
    konto_selected = konto_df.iloc[:, [0, 1, 3, 5, 6, 7, 8]].copy()
    konto_selected.columns = ['Lønnsperiode', 'Ansattnummer', 'Beløp debet', 'Kontostreng', 'Avdeling', 'Project', 'Kampanje']
    
    # Select the required columns from detaljert.xlsx by their actual names
    detaljert_selected = detaljert_df.iloc[:, [3, 4, 5]]
    detaljert_selected.columns = ['Beløp', 'Tekst', 'Reiseregning ID']
    
    # Add a new column 'MacEmp' containing the first three characters of 'Kampanje'
    konto_selected.loc[:, 'Konto'] = konto_selected['Kontostreng'].str.slice(0, 4)
    konto_selected.loc[:, 'Projectno'] = konto_selected['Project'].str.slice(0, 5)
    konto_selected.loc[:, 'Avd'] = konto_selected['Avdeling'].str.split(' ', expand=False).str[0]
    konto_selected.loc[:, 'MacEmp'] = konto_selected['Kampanje'].str.slice(0, 3)
    
    
    
    # Concatenate the selected columns
    combined_df = pd.concat([konto_selected, detaljert_selected], axis=1)
    
    # Write the combined data to a new Excel file
    combined_df.to_excel(output_file, index=False)
    
    # Load the workbook and select the active worksheet
    wb = load_workbook(output_file)
    ws = wb.active
    
    # Define the table range
    table_range = f"A1:{chr(64 + combined_df.shape[1])}{combined_df.shape[0] + 1}"
    
    # Create a table
    tab = Table(displayName="SplicedTable", ref=table_range)
    
    # Add a table style
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    tab.tableStyleInfo = style
    
    # Add the table to the worksheet
    ws.add_table(tab)
    
    # Save the workbook
    wb.save(output_file)
    
    print(f"Data has been written to {output_file}")

if __name__ == "__main__":
    read_and_splice_excel_files()