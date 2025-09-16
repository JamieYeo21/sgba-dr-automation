# Enforcement_data_handler.py
"""
Created on Tuesday 04 March 2025, 08:43:30

Author: Harry Simmons
"""

import pandas as pd
import sys
import os

# Add the adjacent folder to sys.path
folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))
sys.path.append(folder_path)

# Now you can import your functions
from Utility.functions import format_percentage, chop_df, convert_number, get_excel_path
from Utility.dates import sort_dates

def Enforcement_retrieve_data():
    print('Handling Enforcement Data')
    # Accessing the folder which stores the MI tables
    folder_path = 'Q:\\BSP\\Automation\\DR Automation\\Excel_inputs\\[PUT MI TABLES HERE]'
    MI_tables_path = get_excel_path(folder_path)

    # Accessing and Enforcement_1
    Enforcement_1_uncut = pd.read_excel(MI_tables_path, sheet_name='Enforcement_1')
    Enforcement_1 = chop_df(Enforcement_1_uncut, 3, None)
    Enforcement_1 = Enforcement_1[Enforcement_1.iloc[:, 0].str.contains('Total', case=False, na=False)]
    Enforcement_1.reset_index(drop=True, inplace=True)
    Enforcement_1.rename(columns={Enforcement_1.columns[-1]: 'Current Month', Enforcement_1.columns[-2]: 'Last Month'}, inplace=True)
    Enforcement_1['Change'] = Enforcement_1['Current Month'] - Enforcement_1['Last Month']

    # Accessing and Enforcement_2
    Enforcement_2 = pd.read_excel(MI_tables_path, sheet_name='Enforcement_2')
    Enforcement_2 = chop_df(Enforcement_2, 2, None)
    Enforcement_2 = Enforcement_2[Enforcement_2.iloc[:, 0].str.contains('Total', case=False, na=False)]
    Enforcement_2.reset_index(drop=True, inplace=True)

    # Accessing and Enforcement_3
    Enforcement_3 = pd.read_excel(MI_tables_path, sheet_name='Enforcement_3')
    Enforcement_3 = chop_df(Enforcement_3, 2, None)
    Enforcement_3 = Enforcement_3[Enforcement_3.iloc[:, 0].str.contains('Total', case=False, na=False)]
    Enforcement_3.reset_index(drop=True, inplace=True)

    # Accessing and Enforcement_4
    Enforcement_4 = pd.read_excel(MI_tables_path, sheet_name='Enforcement_4')
    Enforcement_4 = chop_df(Enforcement_4, 2, None)
    Enforcement_4 = Enforcement_4[Enforcement_4.iloc[:, 0].str.contains('Total', case=False, na=False)]
    Enforcement_4.reset_index(drop=True, inplace=True)
    Enforcement_4.rename(columns={Enforcement_4.columns[1]: 'Last Month', Enforcement_4.columns[2]: 'Current Month'}, inplace=True)
    Enforcement_4['Change'] = Enforcement_4['Current Month'] - Enforcement_4['Last Month']

    Enforcement_handled_data = {
        'Enforcement_1_uncut': Enforcement_1_uncut,
        'Enforcement_1': Enforcement_1,
        'Enforcement_2': Enforcement_2,
        'Enforcement_3': Enforcement_3,
        'Enforcement_4': Enforcement_4
    }

    print('DONE!')
    return Enforcement_handled_data