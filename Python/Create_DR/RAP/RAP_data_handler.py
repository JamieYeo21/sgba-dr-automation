# RAP_data_handler.py
"""
Created on Tuesday 26 January 2025, 10:39:30

Author: Matthew Bandura
"""

import pandas as pd
import sys
import os

# Add the adjacent folder to sys.path
folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))
sys.path.append(folder_path)

from Utility.functions import chop_df, get_excel_path, format_percentage
from Utility.dates import sort_dates

def RAP_retrieve_data(dates_variables):
    print('Handling RAP Data')
    # Accessing the folder which stores the MI tables
    folder_path = 'Q:\\BSP\\Automation\\DR Automation\\Excel_inputs\\[PUT MI TABLES HERE]'
    MI_tables_path = get_excel_path(folder_path)

    # Accessing the folder which stores the additional stats
    folder_path = 'Q:\\BSP\Automation\\DR Automation\\Excel_inputs\\[PUT ADDITIONAL DR STATS HERE]'
    additional_path = get_excel_path(folder_path)

    # Accessing and transforming Combined_2 (for 11m+)
    Combined_2 = pd.read_excel(MI_tables_path, sheet_name='Combined_2')
    Combined_2 = chop_df(Combined_2, 3, 4)
    Combined_2.rename(columns={Combined_2.columns[1]: '11_18m Number', Combined_2.columns[3]: '18m Number', Combined_2.columns[5]: 'Current Number', Combined_2.columns[-1]: 'Current Percentage'}, inplace=True)
   
    #Accessing and transforming Estimated_2 (for estimates of total no. remediated)
    Estimated_2 = pd.read_excel(MI_tables_path, sheet_name='Estimated_2')
    Estimated_2 = chop_df(Estimated_2, 3, 4)
    Estimated_2.rename(columns = {Estimated_2.columns[1]:'Low Estimate', Estimated_2.columns[2] : 'High Estimate'}, inplace = True)

    #Accessing and transforming Estimated_4 (for estimates of gvt funded remediated) NB. advised to use unrounded numbers (from RAP_misc) instead
    Estimated_4 = pd.read_excel(MI_tables_path, sheet_name='Estimated_4')
    Estimated_4 = chop_df(Estimated_4, 3, 4)
    Estimated_4.rename(columns = {Estimated_4.columns[1]:'Low Estimate', Estimated_4.columns[2] : 'High Estimate'}, inplace = True)


    # Accessing and transforming Combined_8 (for 18m+ in gvt funded schemes)
    Combined_8 = pd.read_excel(MI_tables_path, sheet_name='Combined_8')
    Combined_8 = chop_df(Combined_8, 3, 4)
    Combined_8.rename(columns={Combined_8.columns[1]: '11_18m Number', Combined_8.columns[3]: '18m Number', Combined_8.columns[5]: 'Current Number', Combined_8.columns[-1]: 'Current Percentage'}, inplace=True)

    RAP_handled_data = {
        'Combined_2': Combined_2,
        'Estimated_2': Estimated_2,
        'Estimated_4': Estimated_4,
        'Combined_8': Combined_8,
        'RAP_misc': pd.read_excel(additional_path, sheet_name='RAP_misc')
    }

    print('DONE!')
    return RAP_handled_data
