# ACM_data_handler.py
"""
Created on Monday 16 December 2024, 15:04:03

Author: Harry Simmons
"""

"""
This code 
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
from Utility.functions import format_percentage, chop_df, convert_number, get_excel_path
from Utility.dates import sort_dates

def ACM_retrieve_data(dates_variables, paths_variables):
    print('Handing ACM Data')
    # Imports the date variables
    year = dates_variables['year']

    # Accessing the folder which stores the MI tables
    MI_tables_path = paths_variables['MI_tables_path']

    # Accessing the folder which stores the additional stats
    additional_tables_path = paths_variables['additional_tables_path']

    # Accessing and transforming ACM_2
    ACM_2 = pd.read_excel(MI_tables_path, sheet_name='ACM_2')
    ACM_2 = chop_df(ACM_2, 2, 8)
    ACM_2.rename(columns={ACM_2.columns[-1]: 'Current Percentage', ACM_2.columns[1]: 'Social Number', ACM_2.columns[3]: 'Private Number'}, inplace=True)
    ACM_2['Cumulative Percentage'] = ACM_2['Current Percentage'].cumsum()
    ACM_2.at[7, 'Cumulative Percentage'] = ACM_2.at[7, 'Current Percentage']
    ACM_2['Cumulative Social Percentage'] = ACM_2.iloc[:, 2].cumsum()
    ACM_2['Cumulative Private Percentage'] = ACM_2.iloc[:, 4].cumsum()
    ACM_2['Current'] = ACM_2['Current Percentage'].apply(format_percentage)
    ACM_2['Cumulative Percentage'] = ACM_2['Cumulative Percentage'].apply(format_percentage)
    ACM_2['Cumulative Social Percentage'] = ACM_2['Cumulative Social Percentage'].apply(format_percentage)
    ACM_2['Cumulative Private Percentage'] = ACM_2['Cumulative Private Percentage'].apply(format_percentage)
    ACM_2.at[7, 'Current'] = "100%"
    ACM_2.at[7, 'Cumulative Percentage'] = "100%"
    ACM_2.at[6, 'Cumulative Percentage'] = "100%"
    ACM_2.at[7, 'Cumulative Social Percentag'] = "100%"
    ACM_2.at[7, 'Cumulative Private Percentage'] = "100%"

    # Accessing ACM_6
    ACM_6 = pd.read_excel(MI_tables_path, sheet_name='ACM_6')

    # Accessing and transforming ACM_7
    ACM_7 = pd.read_excel(MI_tables_path, sheet_name='ACM_7')
    ACM_7 = chop_df(ACM_7, 3, 6)
    ACM_7.rename(columns={ACM_7.columns[-1]: 'High', ACM_7.columns[-2]: 'Low'}, inplace=True)

    # Accessing and transforming ACM_11a
    ACM_11 = pd.read_excel(MI_tables_path, sheet_name='ACM_11')
    ACM_11 = chop_df(ACM_11, 4, 8)
    ACM_11 = ACM_11.rename(columns={ACM_11.columns[-1]: 'Current Month', ACM_11.columns[-2]: 'Last Month'})
    ACM_11['Cumulative'] = ACM_11['Current Month'].cumsum()
    ACM_11['Change'] = ACM_11['Current Month'] - ACM_11['Last Month']
    ACM_11['Cumulative Change'] = ACM_11['Change'].cumsum()
    ACM_11.at[7, 'Cumulative'] = ACM_11.at[7, 'Current Month']
    ACM_11.at[7, 'Cumulative Change'] = ACM_11.at[7, 'Change']

    # Accessing and transforming ACM_start_trajectory
    ACM_start_trajectory = pd.read_excel(additional_tables_path, sheet_name='ACM_start_trajectory')
    ACM_start_trajectory['StartedDatePlanned'] = pd.to_datetime(ACM_start_trajectory['StartedDatePlanned'])
    ACM_start_trajectory_filled_NaN = ACM_start_trajectory.copy()
    ACM_start_trajectory_filled_NaN['Vacant'] = ACM_start_trajectory_filled_NaN['Vacant'].fillna('N')
    ACM_start_trajectory_nvt = ACM_start_trajectory_filled_NaN[ACM_start_trajectory_filled_NaN['Vacant'] == 'N']

    # Transforming ACM_6 into ACM yearly identification
    years = ['2017 - 2019'] + [str(y) for y in range(2020, year + 1)]
    ACM_yearly_identified = pd.DataFrame({'Year of identification': years})
    second_column_values = [ACM_6.iloc[10, 11]]
    for i in range(2, (len(years) + 1)):
        second_column_values.append(ACM_6.iloc[i * 10, 11] - ACM_6.iloc[(i - 1) * 10, 11])
    ACM_yearly_identified['Number of buildings identified'] = second_column_values
    ACM_yearly_identified['Cumulative number'] = ACM_yearly_identified['Number of buildings identified'].cumsum()

    # Filter ACM yearly identification to include only the years after 2021 and up to the current year
    ACM_yearly_identified_line = ACM_yearly_identified[(ACM_yearly_identified['Year of identification'] != 'Total') & (ACM_yearly_identified['Year of identification'].astype(str).str.isnumeric())]
    ACM_yearly_identified_line = ACM_yearly_identified_line[ACM_yearly_identified_line['Year of identification'].astype(int) > 2021]
    ACM_yearly_identified_line = ACM_yearly_identified_line[ACM_yearly_identified_line['Year of identification'].astype(int) <= year]
    ACM_yearly_identified_line = ACM_yearly_identified_line.drop(columns=['Cumulative number'])
    ACM_yearly_identified_line['Words'] = ACM_yearly_identified_line['Number of buildings identified'].apply(convert_number)

    # Dictitionary for transporting dataframes to the next script
    ACM_handled_data = {
        'ACM_2': ACM_2,
        'ACM_7': ACM_7,
        'ACM_11': ACM_11,
        'ACM_start_trajectory_nvt': ACM_start_trajectory_nvt,
        'ACM_start_trajectory_filled_NaN': ACM_start_trajectory_filled_NaN,
        'ACM_yearly_identified': ACM_yearly_identified,
        'ACM_yearly_identified_line': ACM_yearly_identified_line
    }  

    print('DONE!')
    return ACM_handled_data
