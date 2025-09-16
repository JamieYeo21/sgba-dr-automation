# functions.py
"""
Created on Wednesday 21 May 2025, 13:49:16

Author: Harry Simmons
"""

import os

def chop_df(df, top_rows_to_drop, bottom_rows_to_drop):
    # Drop rows from 0 to num_rows_to_drop - 1
    df = df.drop(range(top_rows_to_drop))

    # Reset the index
    df = df.reset_index(drop=True)

    # Promote the next row to be the headers
    df.columns = df.iloc[0]
    df = df.drop(0).reset_index(drop=True)

    # Drop bottom rows if bottom_rows_to_drop is not None
    if bottom_rows_to_drop is not None:
        df = df.iloc[:bottom_rows_to_drop]

    return df

def get_excel_path(folder_path):
    files = os.listdir(folder_path)
    excel_files = [file for file in files if file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.xlsm')]
    if len(excel_files) == 1:
        excel_path = os.path.join(folder_path, excel_files[0])
        return excel_path
    else:
        raise FileNotFoundError("No Excel file found in the folder.")