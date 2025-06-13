#
# Read the mapping Excel file into a DataFrame. Skip the first row, as it contains the column names.
# This file is currently not used since API is used for mapping data.
import pandas as pd

def get_mapping_data(mp_filename: str) -> pd.DataFrame:
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
        # Drop the third and fifth column from mpdf DataFrame
        mpdf.drop(mpdf.columns[4], axis=1, inplace=True)
        mpdf.drop(mpdf.columns[2], axis=1, inplace=True)

        # Convert Towards and Project_VAT to string and remove decimal part
        mpdf['Towards'] = mpdf['Towards'].astype(str).str.split('.').str[0]
        mpdf['Project_VAT'] = mpdf['Project_VAT'].astype(str).str.split('.').str[0]

        return mpdf
    
if __name__ == "__main__":
    import os
    # Example usage
    mp_filename = "mapping.xlsx"  
    mapping_df = get_mapping_data(mp_filename)
    print("hello")
    if mapping_df is not None:
        print(mapping_df)