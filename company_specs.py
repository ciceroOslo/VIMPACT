import pandas as pd
import os
import warnings

# Function to create CICERO specific debit/crecit transaction for proper accounting practises
def company_specific_transactions(input_df_hldf: pd.DataFrame, input_df_mapping: pd.DataFrame) -> pd.DataFrame:   
    print(f"\033[96mProcessing CICERO specific transactions...\033[0m")

    # The creation of debit/credit entries to reflect invoiced expenses vs. non-invoiced expenses in the general ledger.
    # The new debit/credit entries are applied on invoicable project (Project>30000) 
    # There are special debit/credit entries for Towards2040 projects (listed in the mapping file)
    # The VAT handling is also considered in the new debit/credit entries.
    
    # Extracting the Towards2040 projects (df_towards).
    df_towards = input_df_mapping.iloc[:, [2]].dropna(subset=[input_df_mapping.columns[2]])
    # Extracting the projects with VAT handeling (df_VAT).
    df_VAT = input_df_mapping.iloc[:, [3]].dropna(subset=[input_df_mapping.columns[3]])

    # print(df_towards)   
   
    # print(df_VAT)

    # print(input_df_hldf)

    # print("Hello nurce!")
       
    new_rows = []

    # Ivoicable projects (>30000) 
    # 5000-5998: Debit 4753, Kredit 5999
    # 6000-6998: Debit 4753, Kredit 6XXX
    # 7000-7998: Debit 4753, Kredit 7999
     
    # Loop through the DataFrame and create the new debit/credit entries
    for index, row in input_df_hldf.iterrows():

        # Deleting VAT codes from non VAT projects
        if (row['Prosjekt']) not in df_VAT['Project_VAT'].values:   
            input_df_hldf.at[index, 'MVA'] = 0 
            row['MVA'] = 0

        # Invoicable projects (not Towards2040 projects)
        # Room for improvement: Use "jobinvoiceable" from Maconomy to identify invoicable projects
        # Values: non-invoiceable, invoiceable, internal_job, internal_job_invoiceable
        # astype(int).values
        if (row['Prosjekt']) > "30000" and (row['Prosjekt']) not in df_towards['Towards'].values:
            # Credit trans

            if int(row['Konto']) in range(5000, 5299) or int(row['Konto']) in range(5330, 5549) or int(row['Konto']) in range(5600, 5998):
                # Same same but different
                input_df_hldf.at[index, 'Konto'] = 4753                
            elif int(row['Konto']) in range(5300, 5329):
                # Debit debit credit 5300-5329
                # credit first...
                new_row_1 = row.copy()
                new_row_1['Konto'] = 5399
                new_row_1['Beløp'] = -row['Beløp']
                new_rows.append(new_row_1)

                # Debit trans on account 4755
                new_row_2 = row.copy()
                new_row_2['Konto'] = 4755
                new_row_2['Beløp'] = abs(row['Beløp'])
                new_rows.append(new_row_2)
            elif int(row['Konto']) in range(5550, 5598):
                # Debit debit credit 5550-5598 (not 5599)
                # credit first...
                new_row_1 = row.copy()
                new_row_1['Konto'] = 5599
                new_row_1['Beløp'] = -row['Beløp']
                new_rows.append(new_row_1)

                # Debit trans on account 4755
                new_row_2 = row.copy()
                new_row_2['Konto'] = 4755
                new_row_2['Beløp'] = abs(row['Beløp'])
                new_rows.append(new_row_2)
            elif int(row['Konto']) in range(6000, 6998):
                 input_df_hldf.at[index, 'Konto'] = 4756 
                
            elif int(row['Konto']) in range(7000, 7169):
                # Debit debit credit
                # credit first...
                new_row_1 = row.copy()
                new_row_1['Konto'] = 7199
                new_row_1['Beløp'] = -row['Beløp']
                new_rows.append(new_row_1)    
     
                # Debit trans on account 4757
                new_row_2 = row.copy()
                new_row_2['Konto'] = 4757
                new_row_2['Beløp'] = abs(row['Beløp'])
                new_rows.append(new_row_2)
            elif int(row['Konto']) in range(7170, 7998):
                # Same same but different
                input_df_hldf.at[index, 'Konto'] = 4757 
         
       
        # Insert the new rows into hldf
        if new_rows:
            df_addded_transactions = pd.concat([input_df_hldf, pd.DataFrame(new_rows)], ignore_index=True)
        
    return df_addded_transactions