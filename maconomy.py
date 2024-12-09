import pandas as pd
import os
import warnings

# Data transformation to Maconomy import file format.
def transform_to_maconomy(input_df_hldf: pd.DataFrame) -> pd.DataFrame:
    print(f"\033[93mTransforming the data to Maconomy format...\033[0m")

    # Building the Maconomy import file from the hldf DataFrame.
    df_macloc = pd.DataFrame(columns=['GeneralJournal:Format','TransactionNumber', 'EntryDate', 'EntryText', 'TypeOfEntry', 'AccountNumber', 'FinanceVATCode', 'DebitBase', 'CreditBase','EntityName','JobNumber','TaskName','ActivityNumber','EmployeeNumber'])    
    df_macloc['EntryDate'] = input_df_hldf['Dato'].dt.strftime('%d/%m/%Y')
    df_macloc['EntryText'] = input_df_hldf.apply(lambda row: row['Text'] if isinstance(row['Text'], str) else None, axis=1)
    df_macloc['TypeOfEntry'] = 'G'
    
    df_macloc['AccountNumber'] = input_df_hldf.apply(lambda row: row['Konto'] if row['Prosjekt'] =="" else None, axis=1)
    df_macloc['FinanceVATCode'] = input_df_hldf.apply(lambda row: row['MVA'] if row['MVA'] > 0 else None, axis=1)
    df_macloc['DebitBase'] = input_df_hldf.apply(lambda row: row['Beløp'] if row['Beløp'] > 0 else None, axis=1)
    df_macloc['CreditBase'] = input_df_hldf.apply(lambda row: abs(row['Beløp']) if row['Beløp'] < 0 else None, axis=1)
    df_macloc['EntityName'] = input_df_hldf.apply(lambda row: row['Avdeling'] if isinstance(row['Avdeling'], str) else None, axis=1)
    df_macloc['ActivityNumber'] = input_df_hldf.apply(lambda row: row['Konto'] if row['Prosjekt'] > "1" else None, axis=1)
    df_macloc['JobNumber'] = input_df_hldf.apply(lambda row: row['Prosjekt'] if row['Prosjekt'] > "0" else None, axis=1)
    df_macloc['EmployeeNumber'] = input_df_hldf.apply(lambda row: row['Medarbeider'] if row['Medarbeider'] != '0' else None, axis=1)
    df_macloc['TaskName'] = input_df_hldf.apply(lambda row: row['Oppgave'] if row['Oppgave'] else None, axis=1)
    df_macloc['GeneralJournal:Format'] = 'GENERALJOURNAL:CREATE'
    df_macloc['TransactionNumber'] = '#KEEP'

    return df_macloc
