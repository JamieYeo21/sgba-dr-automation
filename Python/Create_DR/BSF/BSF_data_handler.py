# BSF_data_handler.py
"""
Created on Thursday 09 January 2025, 10:38:44

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
from Utility.functions import format_percentage, chop_df, get_excel_path
from Utility.dates import sort_dates

def BSF_retrieve_data(paths_variables):
    print('Handing BSF Data')
    # Accessing the folder which stores the MI tables
    MI_tables_path = paths_variables['MI_tables_path']

    # Accessing the folder which stores the additional stats
    additional_tables_path = paths_variables['additional_tables_path']

    # Accessing and transforming BSF_1
    BSF_1 = pd.read_excel(MI_tables_path, sheet_name='BSF_1')
    BSF_1 = chop_df(BSF_1, 4, 6)
    BSF_1.rename(columns={BSF_1.columns[-1]: 'Current Percentage'}, inplace=True)
    BSF_1['Cumulative Percentage'] = BSF_1['Current Percentage'].cumsum()
    BSF_1['Cumulative Social Percentage'] = BSF_1.iloc[:, 6].cumsum()
    BSF_1['Cumulative Private Percentage'] = BSF_1.iloc[:, 8].cumsum()
    BSF_1['Current Percentage'] = BSF_1['Current Percentage'].apply(format_percentage)
    BSF_1['Cumulative Percentage'] = BSF_1['Cumulative Percentage'].apply(format_percentage)
    BSF_1['Cumulative Social Percentage'] = BSF_1['Cumulative Social Percentage'].apply(format_percentage)
    BSF_1['Cumulative Private Percentage'] = BSF_1['Cumulative Private Percentage'].apply(format_percentage)
    BSF_1.at[5, 'Current Percentage'] = "100%"
    BSF_1.at[5, 'Cumulative Percentage'] = "100%"
    BSF_1.at[4, 'Cumulative Percentage'] = "100%"
    BSF_1.at[5, 'Cumulative Social Percentage'] = "100%"
    BSF_1.at[5, 'Cumulative Private Percentage'] = "100%"

    # Accessing and transforming BSF_5a
    BSF_5 = pd.read_excel(MI_tables_path, sheet_name='BSF_5')
    BSF_5 = chop_df(BSF_5, 5, 6)
    BSF_5.rename(columns={BSF_5.columns[0]: 'Remediation Category', BSF_5.columns[-1]: 'Current Month', BSF_5.columns[-2]: 'Last Month', BSF_5.columns[-13]: 'Last Year'}, inplace=True)
    BSF_5['Cumulative'] = BSF_5['Current Month'].cumsum()
    BSF_5['Monthly Change'] = BSF_5['Current Month'] - BSF_5['Last Month']
    BSF_5['Yearly Change'] = BSF_5['Current Month'] - BSF_5['Last Year']
    BSF_5['Cumulative Monthly Change'] = BSF_5['Monthly Change'].cumsum()
    BSF_5['Cumulative Yearly Change'] = BSF_5['Yearly Change'].cumsum()
    BSF_5.at[5, 'Cumulative'] = BSF_5.at[5, 'Current Month']
    BSF_5.at[5, 'Cumulative Monthly Change'] = BSF_5.at[5, 'Monthly Change']
    BSF_5.at[5, 'Cumulative Yearly Change'] = BSF_5.at[5, 'Yearly Change']

    BSF_handled_data = {
        'BSF_1': BSF_1,
        'BSF_5': BSF_5,
        'BSF_reg_status': pd.read_excel(additional_tables_path, sheet_name='BSF_reg_status'),
        'BSF_misc': pd.read_excel(additional_tables_path, sheet_name='BSF_misc'),
    }  

    print('DONE!')
    return BSF_handled_data
