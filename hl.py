import pandas as pd
import os
from datetime import datetime
# import openpyxl



def read_file_to_table():
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    hl_filename = os.path.join(downloads_dir, "HLTrans_971274190_202411.HLT")
    dr_filename = os.path.join(downloads_dir, "Transaksjoner, detaljert.xlsx")
    
    filename = hl_filename
    # filename = input(f"Enter the filename (default is {hl_filename}): ") or hl_filename
    # output_filename = os.path.join(os.path.dirname(filename), "out.xlsx")

    # Defining the column specifications and names based on the fixed-width format in the H & L file
    # ref: https://kundeportal.vismasoftware.no/s/article/Regnskapsformat-Huldt-Lillevik-Standard
    hlcolspecs = [(0, 12), (12, 14), (14, 26), (26, 38), (38, 50), (50, 62), (62, 74), (74, 86), (86, 98), (98, 118), (118, 121), (121, 129), (129, 139), (139, 149), (149, 160) ] 
    hlcolumn_names = ['Konto', 'MVA', 'Avdeling', 'Prosjekt', 'Medarbeider', 'R4', 'R5', 'R6', 'R7', 'ID', 'Filler', 'Dato', 'Ant', 'Sats', 'Beløp']  

    try:
        hldf = pd.read_fwf(filename, colspecs=hlcolspecs, names=hlcolumn_names)
        print("HL-file read successfully into a DataFrame.")
    
    except Exception as e:
        print(f"Error reading the file: {e}")

    try:
        #read the excel file dr_filename into drdf DataFrame
        drdf = pd.read_excel(dr_filename, skiprows=1)
        print("Payroll report Excel file read successfully into a DataFrame.")
    
    except Exception as e:
        print(f"Error reading the file: {e}")    

    # Removing the columns that are not needed
    hldf.drop(columns=['R4', 'R5', 'R6', 'R7', 'Filler', 'Ant', 'Sats'], inplace=True)

    # Converting to proper datatypes - hldf DataFrame
    hldf['Dato'] = pd.to_datetime(hldf['Dato'], format='%d%m%Y', errors='coerce')
    hldf['Beløp'] = hldf['Beløp'].astype(float) / 100
    hldf['Konto'] = pd.to_numeric(hldf['Konto'], errors='coerce')
    hldf['ID'] = pd.to_numeric(hldf['ID'], errors='coerce')

    # If Konto < 5900 and Prosjekt > 0, then Prosjekt = 0
    # Just to clean up wrongly added dimensions from Visma Payroll. This should be done in the payroll system.
    hldf.loc[(hldf['Konto'] < 5900) & (hldf['Prosjekt'] > 0), 'Prosjekt'] = 0
    hldf.loc[(hldf['Medarbeider'] == "0") & (hldf['Prosjekt'] == 0), 'Avdeling'] = 0

    # Use pivot function to aggregate Beløp if Konto and Avdeling and Project are the same
    # This is also to fix the issue where the same transaction is split into multiple lines in the H & L file
    hldf = hldf.pivot_table(index=['Konto', 'Avdeling', 'Prosjekt', 'Medarbeider', 'ID', 'Dato'], values='Beløp', aggfunc='sum').reset_index()

    # And now to the dataframe that contains the payroll report    
    # Drop the columns Lønnsperiode, Ansattnummer og Lønnsart from drdf DataFrame
    drdf.drop(columns=['Lønnsperiode', 'Ansattnummer', 'Lønnsart', 'Beløp'], inplace=True)

    #Use pivot function to aggregate drdf DataFrame if Tekst and Reiseregning ID are the same
    drdf = drdf.pivot_table(index=['Tekst', 'Reiseregning ID']).reset_index()

    #Converting the 'Reiseregning ID' column to a integer object to make merging more robust.
    drdf['Reiseregning ID'] = pd.to_numeric(drdf['Reiseregning ID'], errors='coerce')
    
    # Add the column Text to hldf DataFrame and use a vlookup-like function to get Tekst merged from drdf on IT=Reiseregning ID. 
    hldf['Text'] = hldf['ID'].map(drdf.set_index('Reiseregning ID')['Tekst']).astype(str) 
       
    

    # Just to check the DataFrame
    print(hldf.all)
    print(drdf.all)

    # Write the output to an Excel file
    output_filename = os.path.join(os.path.dirname(filename), "out.xlsx")
    hldf.to_excel(output_filename, index=False)

    check_filename = os.path.join(os.path.dirname(filename), "check.xlsx")
    drdf.to_excel(check_filename, index=False)

    print(f"DataFrame written to {output_filename} successfully.")    


if __name__ == "__main__":
    read_file_to_table()        