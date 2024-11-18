import pandas as pd
import os
import warnings

# Read the H&L file into a DataFrame. Use fixed-width format to read the file.
def process_input_files(hl_filename :str, dr_filename: str, mp_dataframe: pd.DataFrame) -> pd.DataFrame:

    # Mapping - Task lookup from Account.
    mp = mp_dataframe.iloc[:, [0, 1]]

    try:
        # Define the column specifications for the fixed-width Visma Payroll accounting file
        hlcolspecs = [(0, 12), (12, 14), (14, 26), (26, 38), (38, 50), (50, 62), (62, 74), (74, 86), (86, 98), (98, 118), (118, 121), (121, 129), (129, 139), (139, 149), (149, 160), (149,150) ] 
        hlcolumn_names = ['Konto', 'MVA', 'Avdeling', 'Prosjekt', 'Medarbeider', 'R4', 'R5', 'R6', 'R7', 'ID', 'Filler', 'Dato', 'Ant', 'Sats', 'Beløp', 'Sign']  
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
        hldf.loc[(hldf['Konto'] < 5320) | (hldf['Konto'] > 5329) & (hldf['Konto'] < 5849) & (hldf['Prosjekt'] > 0), 'Prosjekt'] = 0
        hldf.loc[(hldf['Medarbeider'] == "0") & (hldf['Prosjekt'] == 0), 'Avdeling'] = 0
        # Populate 'Sign' with "+" if Sign is not "-"
        hldf.loc[(hldf['Sign'] != "-"), 'Sign'] = "+"

        # Populate 'Oppgave' with Task from mp. Mapping should be done on Konto=Account.
        mp.columns = ['Account', 'Task']

        # IF statments to assign a task number if project is specified in the accounting file.     
        hldf.loc[hldf['Prosjekt'] >= 1, 'Oppgave'] = hldf['Konto'].map(mp.set_index('Account')['Task'])
        hldf.loc[hldf['Prosjekt'] < 1, 'Oppgave'] = 0
        
        # Use pivot function to aggregate Beløp if Konto & Avdeling & Project are the same
        # might be removed if Visma Payroll is configured correctly.
        #......need to take into account not to aggregate Beløp with different sign (positive/negative)
      
        hldf = hldf.pivot_table(index=['Konto', 'Avdeling', 'Prosjekt', 'Medarbeider','Oppgave', 'MVA', 'ID', 'Dato', 'Sign'], values='Beløp', aggfunc='sum').reset_index()
        # Don't need the Sign column anymore
        hldf.drop(columns=['Sign'], inplace=True)

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