import pandas as pd
import os
import openpyxl

def process_files(hl_filename, dr_filename):
    try:
        # Define the column specifications for the fixed-width file
        hlcolspecs = [(0, 12), (12, 14), (14, 26), (26, 38), (38, 50), (50, 62), (62, 74), (74, 86), (86, 98), (98, 118), (118, 121), (121, 129), (129, 139), (139, 149), (149, 160) ] 
        hlcolumn_names = ['Konto', 'MVA', 'Avdeling', 'Prosjekt', 'Medarbeider', 'R4', 'R5', 'R6', 'R7', 'ID', 'Filler', 'Dato', 'Ant', 'Sats', 'Beløp']  

        # Read the fixed-width file into a DataFrame
        hldf = pd.read_fwf(hl_filename, colspecs=hlcolspecs, names=hlcolumn_names)
        print("HL-file read successfully into a DataFrame.")
    
    except Exception as e:
        print(f"Error reading the HL file: {e}")
        return None

    try:
        # Read the Excel file into a DataFrame
        drdf = pd.read_excel(dr_filename, skiprows=1)
        print("Payroll report Excel file read successfully into a DataFrame.")
    
    except Exception as e:
        print(f"Error reading the Payroll report file: {e}")
        return None

    try:
        # Removing the columns that are not needed
        hldf.drop(columns=['R4', 'R5', 'R6', 'R7', 'Filler', 'Ant', 'Sats'], inplace=True)

        # Converting to proper datatypes - hldf DataFrame
        hldf['Dato'] = pd.to_datetime(hldf['Dato'], format='%d%m%Y', errors='coerce')
        hldf['Beløp'] = hldf['Beløp'].astype(float) / 100
        hldf['Konto'] = pd.to_numeric(hldf['Konto'], errors='coerce')
        hldf['ID'] = pd.to_numeric(hldf['ID'], errors='coerce')

        # If Konto < 5900 and Prosjekt > 0, then Prosjekt = 0
        hldf.loc[(hldf['Konto'] < 5900) & (hldf['Prosjekt'] > 0), 'Prosjekt'] = 0
        hldf.loc[(hldf['Medarbeider'] == "0") & (hldf['Prosjekt'] == 0), 'Avdeling'] = 0

        # Use pivot function to aggregate Beløp if Konto and Avdeling and Project are the same
        hldf = hldf.pivot_table(index=['Konto', 'Avdeling', 'Prosjekt', 'Medarbeider','MVA', 'ID', 'Dato'], values='Beløp', aggfunc='sum').reset_index()

        # Drop the columns Lønnsperiode, Ansattnummer og Lønnsart from drdf DataFrame
        drdf.drop(columns=['Lønnsperiode', 'Ansattnummer', 'Lønnsart', 'Beløp'], inplace=True)

        # Use pivot function to aggregate drdf DataFrame if Tekst and Reiseregning ID are the same
        drdf = drdf.pivot_table(index=['Tekst', 'Reiseregning ID']).reset_index()

        # Converting the 'Reiseregning ID' column to a integer object to make merging more robust.
        drdf['Reiseregning ID'] = pd.to_numeric(drdf['Reiseregning ID'], errors='coerce')
        
        # Add the column Text to hldf DataFrame and use a vlookup-like function to get Tekst merged from drdf on ID=Reiseregning ID
        hldf['Text'] = hldf['ID'].map(drdf.set_index('Reiseregning ID')['Tekst']).fillna('Lønn ').astype(str)    

    # Insert two rows if "Konto" = 7701
        new_rows = []
        for index, row in hldf.iterrows():
            if row['Konto'] == 7701:
                # Create the first new row with Konto = 7999 and negative Beløp
                new_row_1 = row.copy()
                new_row_1['Konto'] = 7999
                new_row_1['Beløp'] = -row['Beløp']
                new_rows.append(new_row_1)
        
                # Create the second new row with Konto = 4999 and positive Beløp
                new_row_2 = row.copy()
                new_row_2['Konto'] = 4999
                new_row_2['Beløp'] = abs(row['Beløp'])
                new_rows.append(new_row_2)
    
            # Insert the new rows into hldf
            if new_rows:
                hldf = pd.concat([hldf, pd.DataFrame(new_rows)], ignore_index=True)
        
        # Building the Maconomy import file from the hldf DataFrame.
        df_macloc = pd.DataFrame(columns=['GeneralJournal:Format','TransactionNumber', 'EntryDate', 'EntryText', 'TypeOfEntry', 'AccountNumber', 'FinanceVATCode', 'DebitBase', 'CreditBase','EntityName','JobNumber','TaskName','ActivityNumber','EmployeeNumber'])    
        df_macloc['EntryDate'] = hldf['Dato'].dt.strftime('%d/%m/%Y')
        df_macloc['EntryText'] = hldf.apply(lambda row: row['Text'] if isinstance(row['Text'], str) else None, axis=1)
        df_macloc['TypeOfEntry'] = 'G'
        df_macloc['AccountNumber'] = hldf.apply(lambda row: row['Konto'] if row['Prosjekt'] < 1 else None, axis=1)
        df_macloc['FinanceVATCode'] = hldf.apply(lambda row: row['MVA'] if row['MVA'] > 0 else None, axis=1)
        df_macloc['DebitBase'] = hldf.apply(lambda row: row['Beløp'] if row['Beløp'] > 0 else None, axis=1)
        df_macloc['CreditBase'] = hldf.apply(lambda row: abs(row['Beløp']) if row['Beløp'] < 0 else None, axis=1)
        df_macloc['EntityName'] = hldf.apply(lambda row: row['Avdeling'] if isinstance(row['Avdeling'], str) else None, axis=1)
        df_macloc['ActivityNumber'] = hldf.apply(lambda row: row['Konto'] if row['Prosjekt'] > 1 else None, axis=1)
        df_macloc['JobNumber'] = hldf.apply(lambda row: row['Prosjekt'] if row['Prosjekt'] > 0 else None, axis=1)
        df_macloc['EmployeeNumber'] = hldf.apply(lambda row: row['Medarbeider'] if row['Medarbeider'] != '0' else None, axis=1)

        df_macloc['GeneralJournal:Format'] = 'GENERALJOURNAL:CREATE'
        df_macloc['TransactionNumber'] = '#KEEP'
        
        print(df_macloc.all)

        return df_macloc
        #return hldf

    except Exception as e:
        print(f"Error processing the data: {e}")
        return None
    


    except Exception as e:
        print(f"Error writing the data to Maconomy: {e}")

if __name__ == "__main__":
    # Example usage
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    hl_filename = os.path.join(downloads_dir, "HLTrans_971274190_202411.HLT")
    dr_filename = os.path.join(downloads_dir, "Transaksjoner, detaljert.xlsx")

    result_df = process_files(hl_filename, dr_filename)


    if result_df is not None:
        output_filename = os.path.join(os.path.dirname(hl_filename), "out.xlsx")
        result_df.to_excel(output_filename, index=False)
                       
        print(f"DataFrame written to {output_filename} successfully.")