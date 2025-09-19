# Developer_data_handler.py
"""
Created on Wednesday 19 February 2025, 14:39:55

Author: Harry Simmons
"""

import pandas as pd
import sys
import os

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
# Navigate to the Utility folder relative to the script
utility_path = os.path.join(script_dir, 'Utility')
# Add it to sys.path so that python can import from it
sys.path.append(utility_path)

# Now you can import your functions
from Utility.functions import format_percentage, chop_df, convert_number, get_excel_path, create_paths
from Utility.dates import sort_dates

def Developer_retrieve_data_last_month(paths_variables):
    print("Handling Last Month's Developer Data")
    # Accessing the folder which stores the previous month's MI tables
    previous_tables_path = paths_variables['previous_tables_path']

    # Accessing Developer_1
    Developer_1 = pd.read_excel(previous_tables_path, sheet_name='Developer_1')

    # Transforming Developer_1a
    Developer_1a = chop_df(Developer_1, 3, 6)
    Developer_1a.rename(columns={Developer_1a.columns[-2]: 'Current Number'}, inplace=True)
    Developer_1a['Cumulative Number'] = Developer_1a['Current Number'].cumsum()

    # Transforming Developer_1b
    Developer_1b = chop_df(Developer_1, 12, 6)
    Developer_1b.rename(columns={Developer_1b.columns[-2]: 'Current Number'}, inplace=True)
    Developer_1b['Cumulative Number'] = Developer_1b['Current Number'].cumsum()

    # Transforming Developer_1c
    Developer_1c = chop_df(Developer_1, 21, 6)
    Developer_1c.rename(columns={Developer_1c.columns[-2]: 'Current Number'}, inplace=True)
    Developer_1c['Cumulative Number'] = Developer_1c['Current Number'].cumsum()

    # Transforming Developer_1d
    Developer_1d = chop_df(Developer_1, 30, 7)
    Developer_1d.rename(columns={Developer_1d.columns[1]: 'Current Number'}, inplace=True)
    Developer_1d['Cumulative Number'] = Developer_1d['Current Number'].cumsum()

    # Accessing and transforming Developer_2
    Developer_2 = pd.read_excel(previous_tables_path, sheet_name='Developer_2')
    Developer_2 = chop_df(Developer_2, 3, 1)

    # Dictionary for transporting dataframes to the next script
    Developer_handled_data_last_month = {
        'Developer_1a': Developer_1a,
        'Developer_1b': Developer_1b,
        'Developer_1c': Developer_1c,
        'Developer_1d': Developer_1d,
        'Developer_2': Developer_2
    }

    print('DONE!')
    return Developer_handled_data_last_month

def Developer_retrieve_data_this_month(paths_variables):
    print("Handling This Month's Developer Data")
    # Accessing the folder which stores the MI tables
    MI_tables_path = paths_variables['MI_tables_path']

    # Accessing the folder which stores the additional stats
    additional_tables_path = paths_variables['additional_tables_path']


    # Accessing Developer_1
    Developer_1 = pd.read_excel(MI_tables_path, sheet_name='Developer_1')

    # Transforming Developer_1a
    Developer_1a = chop_df(Developer_1, 3, 6)
    Developer_1a.rename(columns={Developer_1a.columns[2]: '11_18m Percentage', Developer_1a.columns[4]: '18m Percentage', Developer_1a.columns[-1]: 'Current Percentage', Developer_1a.columns[-2]: 'Current Number'}, inplace=True)
    Developer_1a['Cumulative Number'] = Developer_1a['Current Number'].cumsum()
    Developer_1a['Cumulative Percentage'] = Developer_1a['Current Percentage'].cumsum()
    Developer_1a['Cumulative 11_18m Percentage'] = Developer_1a['11_18m Percentage'].cumsum()
    Developer_1a['Cumulative 18m Percentage'] = Developer_1a['18m Percentage'].cumsum()
    Developer_1a.at[5, 'Cumulative Number'] = Developer_1a.at[5, 'Current Number']
    Developer_1a['Current'] = Developer_1a['Current Percentage'].apply(format_percentage)
    Developer_1a['Cumulative'] = Developer_1a['Cumulative Percentage'].apply(format_percentage)
    Developer_1a['Cumulative 11_18m Percentage'] = Developer_1a['Cumulative 11_18m Percentage'].apply(format_percentage)
    Developer_1a['Cumulative 18m Percentage'] = Developer_1a['Cumulative 18m Percentage'].apply(format_percentage)
    Developer_1a.at[5, 'Current'] = "100%"
    Developer_1a.at[5, 'Cumulative'] = "100%"
    Developer_1a.at[4, 'Cumulative'] = "100%"

    # Transforming Developer_1b
    Developer_1b = chop_df(Developer_1, 12, 6)
    Developer_1b.rename(columns={Developer_1b.columns[-1]: 'Current Percentage', Developer_1b.columns[-2]: 'Current Number'}, inplace=True)
    Developer_1b['Cumulative Number'] = Developer_1b['Current Number'].cumsum()
    Developer_1b['Cumulative Percentage'] = Developer_1b['Current Percentage'].cumsum()
    Developer_1b['Current'] = Developer_1b['Current Percentage'].apply(format_percentage)
    Developer_1b['Cumulative'] = Developer_1b['Cumulative Percentage'].apply(format_percentage)

    # Transforming Developer_1c
    Developer_1c = chop_df(Developer_1, 21, 6)
    Developer_1c.rename(columns={Developer_1c.columns[-1]: 'Current Percentage', Developer_1c.columns[-2]: 'Current Number'}, inplace=True)
    Developer_1c['Cumulative Number'] = Developer_1c['Current Number'].cumsum()
    Developer_1c['Cumulative Percentage'] = Developer_1c['Current Percentage'].cumsum()
    Developer_1c['Current'] = Developer_1c['Current Percentage'].apply(format_percentage)
    Developer_1c['Cumulative'] = Developer_1c['Cumulative Percentage'].apply(format_percentage)

    # Transforming Developer_1d
    Developer_1d = chop_df(Developer_1, 30, 7)
    Developer_1d.rename(columns={ Developer_1d.columns[2]: 'Current Percentage', Developer_1d.columns[1]: 'Current Number'}, inplace=True)
    Developer_1d['Cumulative Number'] = Developer_1d['Current Number'].cumsum()
    Developer_1d['Cumulative Percentage'] = Developer_1d['Current Percentage'].cumsum()
    Developer_1d['Current Percentage'] = Developer_1d['Current Percentage'].apply(format_percentage)
    Developer_1d['Cumulative Percentage'] = Developer_1d['Cumulative Percentage'].apply(format_percentage)

    # Accessing and transforming Developer_2
    Developer_2 = pd.read_excel(MI_tables_path, sheet_name='Developer_2')
    Developer_2 = chop_df(Developer_2, 3, 1)

    # Accessing and transforming Developer_3
    Developer_3 = chop_df(pd.read_excel(MI_tables_path, sheet_name='Developer_3'), 3, None)

    # Dictionary for transporting dataframes to the next script
    Developer_handled_data_this_month = {
        'Developer_1a': Developer_1a,
        'Developer_1b': Developer_1b,
        'Developer_1c': Developer_1c,
        'Developer_1d': Developer_1d,
        'Developer_2': Developer_2,
        'Developer_3': Developer_3,
        'Developer_misc': pd.read_excel(additional_tables_path, sheet_name='Developer_misc')
    }

    print('DONE!')
    return Developer_handled_data_this_month