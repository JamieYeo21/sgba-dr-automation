# CSS_data_handler.py
"""
Created on Friday 17 January 2025, 14:32:05

Author: Harry Simmons
"""

import pandas as pd
import sys
import os

# Add the adjacent folder to sys.path
folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))
sys.path.append(folder_path)

# Now you can import your functions
from Utility.functions import format_percentage, chop_df, get_excel_path
from Utility.dates import sort_dates

def CSS_retrieve_data(dates_variables):
    print('Handling CSS Data')
    # Accessing the folder which stores the MI tables
    folder_path = 'Q:\\BSP\\Automation\\DR Automation\\Excel_inputs\\[PUT MI TABLES HERE]'
    MI_tables_path = get_excel_path(folder_path)

    # Accessing the folder which stores last month's MI tables
    folder_path = 'Q:\\BSP\\Automation\\DR Automation\\Excel_inputs\\[PUT MI TABLES HERE]\\[LAST MONTHS MI TABLES]'
    old_MI_tables_path = get_excel_path(folder_path)

    # Accessing the folder which stores the additional stats
    folder_path = 'Q:\\BSP\Automation\\DR Automation\\Excel_inputs\\[PUT ADDITIONAL DR STATS HERE]'
    additional_path = get_excel_path(folder_path)

    # Accessing CSS_1 and transforming CSS_1a
    CSS_1 = pd.read_excel(MI_tables_path, sheet_name='CSS_1')
    CSS_1a = chop_df(CSS_1, 3, 3)
    CSS_1a.rename(columns={CSS_1a.columns[1]: 'Number'}, inplace=True)

    # Transforming CSS_1b
    CSS_1b = chop_df(CSS_1, 8, 4)
    CSS_1b.iloc[[0, 2]] = CSS_1b.iloc[[2, 0]].values
    CSS_1b.rename(columns={CSS_1b.columns[-1]: 'Current Percentage'}, inplace=True)
    CSS_1b['Cumulative Percentage'] = CSS_1b['Current Percentage'].cumsum()
    CSS_1b.at[3, 'Cumulative Percentage'] = CSS_1b.at[3, 'Current Percentage']
    CSS_1b['Current Percentage'] = CSS_1b['Current Percentage'].apply(format_percentage)
    CSS_1b['Cumulative Percentage'] = CSS_1b['Cumulative Percentage'].apply(format_percentage)
    CSS_1b.at[3, 'Current Percentage'] = "100%"
    CSS_1b.at[3, 'Cumulative Percentage'] = "100%"
    CSS_1b.at[2, 'Cumulative Percentage'] = "100%"

    # Accessing and transforming CSS_2
    CSS_2 = pd.read_excel(MI_tables_path, sheet_name='CSS_2')
    CSS_2 = chop_df(CSS_2, 3, 11)

    # Accessing and transforming CSS_3
    CSS_3 = pd.read_excel(MI_tables_path, sheet_name='CSS_3')
    CSS_3 = chop_df(CSS_3, 3, 4)
    CSS_3 = CSS_3.rename(columns={CSS_3.columns[-1]: '18m Percentage', CSS_3.columns[-3]: '11_18m Percentage'})
    CSS_3['18m Cumulative Percentage'] = CSS_3['18m Percentage'].cumsum()
    CSS_3['11_18m Cumulative Percentage'] = CSS_3['11_18m Percentage'].cumsum()

    # Accessing and transforming CSS_4
    CSS_4 = pd.read_excel(MI_tables_path, sheet_name='CSS_4')
    CSS_4 = chop_df(CSS_4, 3, 4)
    CSS_4 = CSS_4.rename(columns={CSS_4.columns[-1]: 'Private Percentage', CSS_4.columns[-3]: 'Social Percentage'})
    CSS_4['Private Cumulative Percentage'] = CSS_4['Private Percentage'].cumsum()
    CSS_4['Social Cumulative Percentage'] = CSS_4['Social Percentage'].cumsum()

    # Accessing and transforming CSS_5 (this months')
    CSS_5_current = pd.read_excel(MI_tables_path, sheet_name= 'CSS_5')
    CSS_5_current = chop_df(CSS_5_current, 5, 3)
    CSS_5_current = CSS_5_current.rename(columns={CSS_5_current.columns[1]: 'Private', CSS_5_current.columns[3]: 'Social'})
    CSS_5_current['Total'] = CSS_5_current.loc[1, 'Private'] + CSS_5_current.loc[1, 'Social']
    
    # Accessing and transforming CSS_5 (last months')
    CSS_5_previous = pd.read_excel(old_MI_tables_path, sheet_name= 'CSS_5')
    CSS_5_previous = chop_df(CSS_5_previous, 5, 3)
    CSS_5_previous = CSS_5_previous.rename(columns={CSS_5_previous.columns[1]: 'Private', CSS_5_previous.columns[3]: 'Social'})
    CSS_5_previous['Total'] = CSS_5_previous.loc[1, 'Private'] + CSS_5_previous.loc[1, 'Social']
    

    # Accessing and transforming CSS_6
    CSS_6 = pd.read_excel(MI_tables_path, sheet_name='CSS_6')
    CSS_6 = chop_df(CSS_6, 4, 4)
    CSS_6 = CSS_6.rename(columns={CSS_6.columns[-1]: 'Current Month', CSS_6.columns[-2]: 'Last Month'})
    CSS_6['Cumulative'] = CSS_6['Current Month'].cumsum()
    CSS_6['Change'] = CSS_6['Current Month'] - CSS_6['Last Month']
    CSS_6['Cumulative Change'] = CSS_6['Change'].cumsum()
    CSS_6.at[3, 'Cumulative'] = CSS_6.at[3, 'Current Month']
    CSS_6.at[3, 'Cumulative Change'] = CSS_6.at[3, 'Change']

    # Accessing and transforming CSS_misc
    CSS_misc = pd.read_excel(additional_path, sheet_name='CSS_misc')
    CSS_misc['Change'] = CSS_misc['Current Month'] - CSS_misc['Last Month']

    # Dictionary for transporting dataframes to the next script
    CSS_handled_data = {
        'CSS_1a': CSS_1a,
        'CSS_1b': CSS_1b,
        'CSS_2': CSS_2,
        'CSS_3': CSS_3,
        'CSS_4': CSS_4,
        'CSS_5_current' : CSS_5_current,
        'CSS_5_previous' : CSS_5_previous,
        'CSS_6': CSS_6,
        'CSS_misc': CSS_misc
    }

    print('DONE!')
    return CSS_handled_data
