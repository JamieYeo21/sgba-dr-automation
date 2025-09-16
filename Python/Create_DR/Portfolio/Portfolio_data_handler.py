# Portfolio_data_handler.py
"""
Created on Monday 27 January 2025, 09:14:42

Author: Harry Simmons
"""

import pandas as pd
import sys
import os

# Add the adjacent folder to sys.path
folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))
sys.path.append(folder_path)

from Utility.functions import chop_df, get_excel_path, format_percentage
from Utility.dates import sort_dates

def Portfolio_retrieve_data(date_variables):
    print('Handling Portfolio Data')
    # Accessing the folder which stores the MI tables
    folder_path = 'Q:\\BSP\\Automation\\DR Automation\\Excel_inputs\\[PUT MI TABLES HERE]'
    MI_tables_path = get_excel_path(folder_path)

    # Accessing and transforming Combined_2
    Combined_2 = pd.read_excel(MI_tables_path, sheet_name='Combined_2')
    Combined_2 = chop_df(Combined_2, 3, 4)
    Combined_2.rename(columns={Combined_2.columns[1]: '11_18m Number', Combined_2.columns[3]: '18m Number', Combined_2.columns[5]: 'Current Number', Combined_2.columns[-1]: 'Current Percentage'}, inplace=True)
    Combined_2['Cumulative Percentage'] = Combined_2['Current Percentage'].cumsum()
    Combined_2.at[3, 'Cumulative Percentage'] = Combined_2.at[3, 'Current Percentage']
    Combined_2['Cumulative 11_18m Percentage'] = Combined_2.iloc[:, 2].cumsum()
    Combined_2['Cumulative 18m Percentage'] = Combined_2.iloc[:, 4].cumsum()
    Combined_2['Current Percentage'] = Combined_2['Current Percentage'].apply(format_percentage)
    Combined_2['Cumulative Percentage'] = Combined_2['Cumulative Percentage'].apply(format_percentage)
    Combined_2['Cumulative 11_18m Percentage'] = Combined_2['Cumulative 11_18m Percentage'].apply(format_percentage)
    Combined_2['Cumulative 18m Percentage'] = Combined_2['Cumulative 18m Percentage'].apply(format_percentage)
    Combined_2.at[3, 'Current Percentage'] = "100%"
    Combined_2.at[3, 'Cumulative Percentage'] = "100%"
    Combined_2.at[2, 'Cumulative Percentage'] = "100%"
    Combined_2.at[3, 'Cumulative 11_18m Percentage'] = "100%"
    Combined_2.at[3, 'Cumulative 18m Percentage'] = "100%"

    # Accessing and transforming Combined_4
    Combined_4 = pd.read_excel(MI_tables_path, sheet_name='Combined_4')
    Combined_4 = chop_df(Combined_4, 3, 4)
    Combined_4.rename(columns={Combined_4.columns[-1]: 'Total Dwellings'}, inplace=True)

    # Accessing and transforming Combined_5
    Combined_5 = pd.read_excel(MI_tables_path, sheet_name='Combined_5')
    Combined_5 = chop_df(Combined_5, 3, 4)
    Combined_5.rename(columns={Combined_5.columns[2]: 'Private Percentage', Combined_5.columns[4]: 'Social Percentage'}, inplace=True)
    Combined_5['Cumulative Private Percentage'] = Combined_5.iloc[:, 2].cumsum()
    Combined_5['Cumulative Social Percentage'] = Combined_5.iloc[:, 4].cumsum()
    Combined_5['Cumulative Private Percentage'] = Combined_5['Cumulative Private Percentage'].apply(format_percentage)
    Combined_5['Cumulative Social Percentage'] = Combined_5['Cumulative Social Percentage'].apply(format_percentage)

    # Accessing and transforming Combined_6
    Combined_6 = pd.read_excel(MI_tables_path, sheet_name='Combined_6')
    Combined_6 = chop_df(Combined_6, 4, 4)
    Combined_6.rename(columns={Combined_6.columns[-1]: 'Current Month', Combined_6.columns[-2]: 'Last Month', Combined_6.columns[-13]: 'Last Year', Combined_6.columns[11]: 'October 2023'}, inplace=True)
    Combined_6['Cumulative'] = Combined_6['Current Month'].cumsum()
    Combined_6['Monthly Change'] = Combined_6['Current Month'] - Combined_6['Last Month']
    Combined_6['Yearly Change'] = Combined_6['Current Month'] - Combined_6['Last Year']
    Combined_6['Since October 2023'] = Combined_6['Current Month'] - Combined_6['October 2023']
    Combined_6['Cumulative Monthly Change'] = Combined_6['Monthly Change'].cumsum()
    Combined_6['Cumulative Yearly Change'] = Combined_6['Yearly Change'].cumsum()
    Combined_6.at[3, 'Cumulative'] = Combined_6.at[3, 'Current Month']
    Combined_6.at[3, 'Cumulative Monthly Change'] = Combined_6.at[3, 'Monthly Change']
    Combined_6.at[3, 'Cumulative Yearly Change'] = Combined_6.at[3, 'Yearly Change']

    Portfolio_handled_data = {
        'Combined_2': Combined_2,
        'Combined_4': Combined_4,
        'Combined_5': Combined_5,
        'Combined_6': Combined_6
    }

    print('DONE!')
    return Portfolio_handled_data
