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
    

def create_paths():
    home_dir = os.path.expanduser("~")

    # Build the path to the graph output folder
    figure_path = os.path.join(home_dir, "GitHub", "sgba-DR-automation", "DR_outputs", "DR_graphs")

    previous_MI_folder = os.path.join(home_dir,  "GitHub", "sgba-DR-automation", "Excel_inputs","[PUT MI TABLES HERE]", "[LAST MONTHS MI TABLES]")    
    previous_tables_path = get_excel_path(previous_MI_folder)

    MI_folder = os.path.join(home_dir,  "GitHub", "sgba-DR-automation", "Excel_inputs","[PUT MI TABLES HERE]")    
    MI_tables_path = get_excel_path(MI_folder)

    additional_folder = os.path.join(home_dir,  "GitHub", "sgba-DR-automation", "Excel_inputs","[PUT ADDITIONAL DR STATS HERE]")
    additional_tables_path = get_excel_path(additional_folder)

    partial_output_path = os.path.join(home_dir, "GitHub", "sgba-DR-automation", "DR_outputs", "DR_graphs")

    save_path = os.path.join(home_dir, "GitHub", "sgba-dr-automation", "DR_outputs", "Auto_DR")

    paths = {'figure_path' : figure_path,
             'previous_tables_path' : previous_tables_path,
             'MI_tables_path' : MI_tables_path,
             'additional_tables_path' : additional_tables_path,
             'save_path' : save_path,
             'partial_output_path' : partial_output_path
    }
    return paths

