# Read the mapping Excel file into a DataFrame. Skip the first row, as it contains the column names.
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
        return mpdf