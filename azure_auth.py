# Description: This script is used to authenticate the user and get the access token to access the Maconomy API.
# The last part is fetching the data from the Maconomy API and returning it as a DataFrame.

from msal import PublicClientApplication
import requests
import pandas as pd

def get_mapping_api(client_id: str, tenant_id: str, scopes: list[str], api_gateway: str)-> pd.DataFrame:

    # Defining the Azure and App registration ID values
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    app = PublicClientApplication(client_id, authority=authority)

    # Attempt to get a token silently
    accounts = app.get_accounts()
    result = app.acquire_token_silent(scopes, account=accounts[0]) if accounts else None
    
    # If no token is found, use interactive login
    if not result:
        result = app.acquire_token_interactive(scopes)

    access_token = result["access_token"]
    headers = {"Authorization": f"Bearer {access_token}"}

    # print(access_token)
    api_url_1 = f"{api_gateway}/jobs/filter?fields=jobnumber&restriction=specification4name%20like%20\"Towards2040\"&limit=1000"
    api_url_2 = f"{api_gateway}/jobs/filter?fields=jobname&restriction=vat%20and%20not(closed)%20and%20not(template)&limit=1000"
    api_url_3 = f"{api_gateway}/AccountCard/filter?restriction=statistic3%20>%20\"1\"&orderby=accountnumber&fields=accountnumber, statistic3&limit=1000"

    # Room for improvement: Use "jobinvoiceable" from Maconomy to identify invoicable projects
    # Values: non-invoiceable, invoiceable, internal_job, internal_job_invoiceable

    toward = requests.get(api_url_1, headers=headers)
    vat = requests.get(api_url_2, headers=headers)
    task = requests.get(api_url_3, headers=headers)   

    towards = toward.json()
    vats = vat.json()
    tasks = task.json()

    # Towards 2024 projects
    if 'panes' in towards and 'filter' in towards['panes'] and 'records' in towards['panes']['filter']:
        records = towards['panes']['filter']['records']
        rows = [{'jobnumber': record['data']['jobnumber']} for record in records]
    else:
        rows = []

    towards_df = pd.DataFrame(rows, columns=['jobnumber'])
    # towards_df = pd.DataFrame(rows, columns=['jobnumber', 'specification4name'])

    # Projects with VAT
    if 'panes' in vats and 'filter' in vats['panes'] and 'records' in vats['panes']['filter']:
        records = vats['panes']['filter']['records']
        rows = [{'jobnumber': record['data']['jobnumber']} for record in records]
    else:
        rows = []

    vats_df = pd.DataFrame(rows, columns=['jobnumber'])

    # Account to task number mapping
    if 'panes' in tasks and 'filter' in tasks['panes'] and 'records' in tasks['panes']['filter']:
        records = tasks['panes']['filter']['records']
        rows = [{'accountnumber': record['data']['accountnumber'], 'statistic3': record['data']['statistic3']} for record in records]
    else:
        rows = []

    tasks_df = pd.DataFrame(rows, columns=['accountnumber', 'statistic3'])

    # Stack the coloumns of the dataframes in one dataframe with the same index
    df = pd.concat([tasks_df, towards_df, vats_df], axis=1)
    # rename df columns
    df.columns = ['Account', 'Task', 'Towards',  'Project_VAT']

    return df

if __name__ == "__main__":
    # Just for testing purposes...
    client_id = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"  
    tenant_id = "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"
    scopes = ["api://zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz/.default"] 
    api_gateway = "https://xyz.azure-api.net/mac"
    mapping_df = get_mapping_api(client_id, tenant_id, scopes, api_gateway)
    print(mapping_df)