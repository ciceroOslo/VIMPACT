# VIMPACT by Frode Rørvik, CICERO Center for International Climate Research 
# Date: 2024-12-12

import pandas as pd
import os
import warnings
# Importing the functions from the modules
from preprosessing import process_input_files
from company_specs import company_specific_transactions
from maconomy import transform_to_maconomy
from datetime import datetime, timedelta

# Choose API or Excel for mapping data
from azure_auth import get_mapping_api
# from get_mapping import get_mapping_data

# Debugging help - print all rows in the DataFrame
# pd.set_option('display.max_rows', None)

# ***********************************************************************************
# The main program code                                                             *
# ***********************************************************************************        

def main() -> None:
    # Assigning the organization number for the Payroll file name
    orgno:          str = "971274190"
    
    # Define the directory where the files are stored (users download directory)
    downloads_dir:  str = os.path.join(os.path.expanduser("~"), "Downloads")

    # API and Ouauth2.0 authentication
    # We are using Azure APIM as a gateway to Maconomy and Entra ID for authentication (user auth)
	client_id = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"  # Application (client) ID of app registration
    tenant_id = "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" # Directory (tenant) ID of tenant
    scopes = ["api://zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz/.default"] # The clientID of the API app registration
    api_gateway = "https://xyz.azure-api.net/mac"

    # Calculate the date part of the accounting file name
    today = datetime.today()
    first_day_of_month = today.replace(day=1)
    datepart:       str = first_day_of_month.strftime("%Y%m")
    # Define the input files   
    hl_filename:    str = os.path.join(downloads_dir, "HLTrans_" + orgno + "_" + datepart + ".HLT")
    dr_filename:    str = os.path.join(downloads_dir, "Transaksjoner, detaljert.xlsx")

    # Getting the mapping data from the Excel file or API
    # mp_filename:    str = os.path.join("mapping.xlsx")
    # mapping_df: pd.DataFrame = get_mapping_data(mp_filename)
    mapping_df: pd.DataFrame = get_mapping_api(client_id, tenant_id, scopes, api_gateway)

    # Processing and preparing the accounting data
    accounting_df: pd.DataFrame = process_input_files(hl_filename, dr_filename, mapping_df)

    # Adding CICERO specific debit/credit transactions to the accounting data
    cicero_accounting_df: pd.DataFrame = company_specific_transactions(accounting_df, mapping_df)

    # Transforming the accounting data to Maconomy format.
    # NB: If you do not want the CICERO-sepcific transactions, you can specify accounting_df instead of cicero_accounting_df.
    maconomy_df: pd.DataFrame = transform_to_maconomy(cicero_accounting_df)

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
        except Exception as e:
            print(f"\033[91mError writing the Maconomy import file: {e}\033[0m")
        else:                      
            print(f"\033[95mDataFrame written to {output_filename} successfully.\033[0m")

if __name__ == "__main__":
    main()

### Fine