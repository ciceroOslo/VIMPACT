# VIMPACT by Frode Rørvik, CICERO Center for International Climate Research 
# Date: 2024-10-20


import pandas as pd
import os
import warnings

# Read the mapping Excel file into a DataFrame. Skip the first row, as it contains the column names.
def get_mapping_data(mp_filename):
    try:
        mpdf = pd.read_excel(mp_filename)   
    except FileExistsError as e:
        print(f"\033[91mError: The file {e} is in use by another application or file not found.\033[0m")
        return None
    except Exception as e:
        print(f"\033[91mError reading the file. The file might be in use: {e}\033[0m")
        return None
    else:
        print(f"\033[92m1) The mapping Excel file was read successfully.\033[0m")
        return mpdf

# Read the H&L file into a DataFrame. Use fixed-width format to read the file.
def process_input_files(hl_filename, dr_filename, mp_dataframe):

    # Mapping - Task lookup from Account.
    mp = mp_dataframe.iloc[:, [0, 2]]

    try:
        # Define the column specifications for the fixed-width Visma Payroll accounting file
        hlcolspecs = [(0, 12), (12, 14), (14, 26), (26, 38), (38, 50), (50, 62), (62, 74), (74, 86), (86, 98), (98, 118), (118, 121), (121, 129), (129, 139), (139, 149), (149, 160) ] 
        hlcolumn_names = ['Konto', 'MVA', 'Avdeling', 'Prosjekt', 'Medarbeider', 'R4', 'R5', 'R6', 'R7', 'ID', 'Filler', 'Dato', 'Ant', 'Sats', 'Beløp']  
        # Read the fixed-width file into a DataFrame
        hldf = pd.read_fwf(hl_filename, colspecs=hlcolspecs, names=hlcolumn_names)
    except FileExistsError as e:
        print(f"\033[91mError: The file {e} is in use by another application or file not found.\033[0m")
        return None
    except Exception as e:
        print(f"\033[91mError reading the file. The file might be in use: {e}\033[0m")
        return None
    else:    
        print(f"\033[92m2) The HL Payroll accounting file read successfully.\033[0m")

    # Read the supporting Visma Payroll report file (Excel) - Transaksjoner, detaljert.xlsx into a DataFrame.
    try:
        warnings.filterwarnings("ignore", category=UserWarning, message="Workbook contains no default style, apply openpyxl's default")    
        drdf = pd.read_excel(dr_filename, skiprows=1, engine='openpyxl')
    except FileExistsError as e:
        print(f"\033[91mError: The file {e} is in use by another application or file not found.\033[0m")
        return None
    except Exception as e:
        print(f"\033[91mError reading the file. The file might be in use: {e}\033[0m")
        return None
    else:
        print(f"\033[92m3) Payroll report Excel file read successfully.\033[0m")

    #########################################################################################################
    # Data processing and transformation
    #########################################################################################################
    try:
        # Removing the columns that are not needed in the Payroll accounting dataframe
        hldf.drop(columns=['R4', 'R5', 'R6', 'R7', 'Filler', 'Ant', 'Sats'], inplace=True)

        # Insert column 'Oppgave' between 'Prosjekt' and 'Medarbeider'
        hldf.insert(hldf.columns.get_loc('Medarbeider'), 'Oppgave', None)

        # Converting to proper datatypes - hldf DataFrame
        hldf['Dato'] = pd.to_datetime(hldf['Dato'], format='%d%m%Y', errors='coerce')
        hldf['Beløp'] = hldf['Beløp'].astype(float) / 100
        hldf['Konto'] = pd.to_numeric(hldf['Konto'], errors='coerce')
        hldf['ID'] = pd.to_numeric(hldf['ID'], errors='coerce')

        # Removing unwanted Project and Department dimensions. Just to correct sub-optimal configuration of Visma Payroll.
        # This might be removed in the future.
        hldf.loc[(hldf['Konto'] < 5320) | (hldf['Konto'] > 5329) & (hldf['Konto'] < 5800) & (hldf['Prosjekt'] > 0), 'Prosjekt'] = 0
        hldf.loc[(hldf['Medarbeider'] == "0") & (hldf['Prosjekt'] == 0), 'Avdeling'] = 0

        # Populate 'Oppgave' with Task from mp. Mapping should be done on Konto=Account.
        mp.columns = ['Account', 'Task']

        # If Prosjekt > 10000 then map Oppgave from Konto using the mapping DataFrame else Oppgave = 0        
        hldf.loc[hldf['Prosjekt'] > 10000, 'Oppgave'] = hldf['Konto'].map(mp.set_index('Account')['Task'])
        hldf.loc[hldf['Prosjekt'] < 10000, 'Oppgave'] = 0
        
        # Use pivot function to aggregate Beløp if Konto & Avdeling & Project are the same
        # might be removed in the future.
        hldf = hldf.pivot_table(index=['Konto', 'Avdeling', 'Prosjekt', 'Medarbeider','Oppgave', 'MVA', 'ID', 'Dato'], values='Beløp', aggfunc='sum').reset_index()

        # Drop the columns Lønnsperiode, Ansattnummer og Lønnsart from drdf DataFrame (Payroll Excel report)
        drdf.drop(columns=['Lønnsperiode', 'Ansattnummer', 'Lønnsart', 'Beløp', 'MVA-kode'], inplace=True)

        # Use pivot function to aggregate drdf if rows are duplicated (Tekst & Reiseregning ID)
        drdf = drdf.pivot_table(index=['Tekst', 'Reiseregning ID']).reset_index()

        # Converting the 'Reiseregning ID' to an integer object to make merging/mapping more robust.
        drdf['Reiseregning ID'] = pd.to_numeric(drdf['Reiseregning ID'], errors='coerce')
        
        # Merging: Add the column Text to hldf DataFrame and use a vlookup-like function to fetch drdf and join on ID=Reiseregning ID     
        hldf['Text'] = hldf.apply(lambda row: f"{drdf.set_index('Reiseregning ID')['Tekst'].get(row['ID'], 
        f'Lønn ({row['Dato'].strftime('%Y-%m-%d')})')}" 
        if 'Lønn' in drdf.set_index('Reiseregning ID')['Tekst'].get(row['ID'], 
        f'Lønn ({row['Dato'].strftime('%Y-%m-%d')})') 
        else f"{drdf.set_index('Reiseregning ID')['Tekst'].get(row['ID'], 
        f'Lønn ({row['Dato'].strftime('%Y-%m-%d')})')} ({row['ID']})", axis=1)

    except Exception as e:
        print(f"\033[91mError processing the data: {e}\033[0m")
        return None
    finally:
        #return hldf dataframe from the function
        
        print(f"\033[94mRescuing Visma from the clutches of product manager misadventures...\033[0m")
        return hldf
    

# Function to create CICERO specific debit/crecit transaction for proper accounting practises
def cicero_specific_transactions(input_df_hldf, input_df_mapping):   
    print(f"\033[96mProcessing CICERO specific transactions...\033[0m")

    # The creation of debit/credit entries to reflect invoiced expenses vs. non-invoiced expenses in the general ledger.
    # The new debit/credit entries are applied on invoicable project (Project>30000) 
    # There are special debit/credit entries for Towards2040 projects (listed in the mapping file)
    # The VAT handling is also considered in the new debit/credit entries.
    
    # Extracting the Towards2040 projects (df_towards).
    df_towards = input_df_mapping.iloc[:, [4]].dropna(subset=[input_df_mapping.columns[4]])
    # Extracting the projects with VAT handeling (df_VAT).
    df_VAT = input_df_mapping.iloc[:, [6]].dropna(subset=[input_df_mapping.columns[6]])
       
    new_rows = []

    # Ivoicable projects (>30000) 
    # 5000-5998: Debit 4753, Kredit 5999
    # 6000-6998: Debit 4753, Kredit 6XXX
    # 7000-7998: Debit 4753, Kredit 7999

    # Towards (df_towards)
    # 5000-5998: Debit 7713, Kredit 5999
    # 6000-6998: Debit 7713, Kredit 6XXX
    # 7000-7998: Debit 7713, Kredit 7999
    
    # Loop through the DataFrame and create the new debit/credit entries
    for index, row in input_df_hldf.iterrows():

        # Deleting VAT codes from non VAT projects
        if int(row['Prosjekt']) not in df_VAT['Project_VAT'].astype(int).values:   
            input_df_hldf.at[index, 'MVA'] = 0 
            row['MVA'] = 0

        if int(row['Prosjekt']) > 30000 and int(row['Prosjekt']) not in df_towards['Towards'].astype(int).values:
            # Ordenary invoiable project - credit trans.
            if int(row['Konto']) in range(5000, 5999):
                new_row_1 = row.copy()
                new_row_1['Konto'] = 5999
                new_row_1['Beløp'] = -row['Beløp']
                new_rows.append(new_row_1)
            elif int(row['Konto']) in range(6000, 6999):
                new_row_1 = row.copy()
                new_row_1['Konto'] = ['Konto']
                new_row_1['Beløp'] = -row['Beløp']
                new_rows.append(new_row_1)
            elif int(row['Konto']) in range(7000, 7999):
                new_row_1 = row.copy()
                new_row_1['Konto'] = 7999
                new_row_1['Beløp'] = -row['Beløp']
                new_rows.append(new_row_1)    
     
            # Debit trans on account 4753
            new_row_2 = row.copy()
            new_row_2['Konto'] = 4753
            new_row_2['Beløp'] = abs(row['Beløp'])
            new_rows.append(new_row_2)

        elif int(row['Prosjekt']) in df_towards['Towards'].astype(int).values:
            # Towards2040 projects (df_towards) - credit trans.
            if int(row['Konto']) in range(5000, 5999):
                new_row_1 = row.copy()
                new_row_1['Konto'] = 5999
                new_row_1['Beløp'] = -row['Beløp']
                new_rows.append(new_row_1)
            elif int(row['Konto']) in range(6000, 6999):
                new_row_1 = row.copy()
                new_row_1['Konto'] = ['Konto']
                new_row_1['Beløp'] = -row['Beløp']
                new_rows.append(new_row_1)
            elif int(row['Konto']) in range(7000, 7999):
                new_row_1 = row.copy()
                new_row_1['Konto'] = 7999
                new_row_1['Beløp'] = -row['Beløp']
                new_rows.append(new_row_1)    
             
            # Debit trans on account 7713
            new_row_2 = row.copy()
            new_row_2['Konto'] = 7713
            new_row_2['Beløp'] = abs(row['Beløp'])
            new_rows.append(new_row_2)
                   
        # Insert the new rows into hldf
        if new_rows:
            df_addded_transactions = pd.concat([input_df_hldf, pd.DataFrame(new_rows)], ignore_index=True)
        
    return df_addded_transactions

# Data transformation to Maconomy import file format.
def transform_to_maconomy(input_df_hldf):
    print(f"\033[93mTransforming the data to Maconomy format...\033[0m")

    # Building the Maconomy import file from the hldf DataFrame.
    df_macloc = pd.DataFrame(columns=['GeneralJournal:Format','TransactionNumber', 'EntryDate', 'EntryText', 'TypeOfEntry', 'AccountNumber', 'FinanceVATCode', 'DebitBase', 'CreditBase','EntityName','JobNumber','TaskName','ActivityNumber','EmployeeNumber'])    
    df_macloc['EntryDate'] = input_df_hldf['Dato'].dt.strftime('%d/%m/%Y')
    df_macloc['EntryText'] = input_df_hldf.apply(lambda row: row['Text'] if isinstance(row['Text'], str) else None, axis=1)
    df_macloc['TypeOfEntry'] = 'G'
    df_macloc['AccountNumber'] = input_df_hldf.apply(lambda row: row['Konto'] if row['Prosjekt'] < 1 else None, axis=1)
    df_macloc['FinanceVATCode'] = input_df_hldf.apply(lambda row: row['MVA'] if row['MVA'] > 0 else None, axis=1)
    df_macloc['DebitBase'] = input_df_hldf.apply(lambda row: row['Beløp'] if row['Beløp'] > 0 else None, axis=1)
    df_macloc['CreditBase'] = input_df_hldf.apply(lambda row: abs(row['Beløp']) if row['Beløp'] < 0 else None, axis=1)
    df_macloc['EntityName'] = input_df_hldf.apply(lambda row: row['Avdeling'] if isinstance(row['Avdeling'], str) else None, axis=1)
    df_macloc['ActivityNumber'] = input_df_hldf.apply(lambda row: row['Konto'] if row['Prosjekt'] > 1 else None, axis=1)
    df_macloc['JobNumber'] = input_df_hldf.apply(lambda row: row['Prosjekt'] if row['Prosjekt'] > 0 else None, axis=1)
    df_macloc['EmployeeNumber'] = input_df_hldf.apply(lambda row: row['Medarbeider'] if row['Medarbeider'] != '0' else None, axis=1)
    df_macloc['TaskName'] = input_df_hldf.apply(lambda row: row['Oppgave'] if row['Oppgave'] else None, axis=1)
    df_macloc['GeneralJournal:Format'] = 'GENERALJOURNAL:CREATE'
    df_macloc['TransactionNumber'] = '#KEEP'
    return df_macloc

# ***********************************************************************************
# The main program code - make sure it only runs if the script is executed directly *
# ***********************************************************************************        

if __name__ == "__main__":
    # Define the file paths and input files
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    hl_filename = os.path.join(downloads_dir, "HLTrans_971274190_202411.HLT")
    dr_filename = os.path.join(downloads_dir, "Transaksjoner, detaljert.xlsx")
    mp_filename = os.path.join("mapping.xlsx")

    # Getting the mapping data from the Excel file
    mapping_df = get_mapping_data(mp_filename)

    # Processing and preparing the accounting data
    accounting_df = process_input_files(hl_filename, dr_filename, mapping_df)

    # Adding CICERO specific debit/credit transactions to the accounting data
    cicero_accounting_df = cicero_specific_transactions(accounting_df, mapping_df)

    # Transforming the accounting data to Maconomy format.
    # NB: If you do not want the CICERO-sepcific transactions, you can specify accounting_df instead of cicero_accounting_df.
    maconomy_df = transform_to_maconomy(cicero_accounting_df)

    # Static assignment of a DataFrame with three columns
    mac_header_df = pd.DataFrame({
        'Column1': ['JOURNAL:Format','JOURNAL:CREATE'],
        'Column2': ['TransactionNumberSeries','Lønn'],
        'Column3': ['CompanyNumber','1']
    })

    # Writing the Maconomy DataFrame to an Excel file
    if maconomy_df is not None:
        output_filename = os.path.join(os.path.dirname(hl_filename), "out.xlsx")

        try:
            mac_header_df.to_excel(output_filename, index=False, header=False)
            with pd.ExcelWriter(output_filename, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                maconomy_df.to_excel(writer, index=False, startrow=len(mac_header_df) + 1)

           # maconomy_df.to_excel(output_filename, index=False)
        except Exception as e:
            print(f"\033[91mError writing the Maconomy import file: {e}\033[0m")
        else:                      
            print(f"\033[95mDataFrame written to {output_filename} successfully.\033[0m")

### Fine