# ACM_variables.py
"""
Created on Monday 16 December 2024, 15:31:08

Author: Harry Simmons
"""

from datetime import datetime
import calendar as cal
import pandas as pd
import sys
import os

# Add the Utility folder to sys.path
folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))
sys.path.append(folder_path)

from ACM.ACM_data_handler import ACM_retrieve_data
from Utility.functions import Change_line_in_DR, format_percentage, convert_number
from Utility.dates import sort_dates

def ACM_variable_creator(ACM_handled_data, dates_variables):
    # Imports the date variables
    end_quarter_no = dates_variables['end_quarter_no']
    end_quarter_word = dates_variables['end_quarter_word']
    year = dates_variables['year']
    end_this_year = dates_variables['end_this_year']
    end_next_year = dates_variables['end_next_year']

    # Unpacking dataframes from ACM_data_handler
    ACM_2 = ACM_handled_data['ACM_2']
    ACM_7 = ACM_handled_data['ACM_7']
    ACM_11 = ACM_handled_data['ACM_11']
    ACM_start_trajectory_nvt = ACM_handled_data['ACM_start_trajectory_nvt']
    ACM_start_trajectory_filled_NaN = ACM_handled_data['ACM_start_trajectory_filled_NaN']
    ACM_yearly_identified = ACM_handled_data['ACM_yearly_identified']
    ACM_yearly_identified_line = ACM_handled_data['ACM_yearly_identified_line']

    # Formatting the dataframe for the ACM remediation table
    ACM_remediation_table = pd.DataFrame({
        'Remediation Stage' : ['Completed Remediation', 'Remediation complete awaiting building control signoff', 'Remediation started - cladding removed', 'Remediation started', 'Remediation plans in place', 'Intent to remediate', 'Remediation plan unclear', 'Total'],
        'Number of buildings' : ACM_11['Current Month'],
        'Percentage' : ACM_2['Current'],
        'Cumulative Number' : ACM_11['Cumulative'],
        'Cumulative Percentage' : ACM_2['Cumulative Percentage']
    })

    # Formatting the dataframe for the ACM enforcement table
    ACM_enforcement_end_quarter_line = f'Forecast to start by the end of {end_quarter_word}'
    ACM_enforcement_table_data = [
        ['Total', (ACM_start_trajectory_nvt['Enforcement'] == 'Yes').sum(), (ACM_start_trajectory_nvt['JIT'] == 'Yes').sum(), (ACM_start_trajectory_nvt['StartedDatePlanned'] > datetime(2000, 1, 1)).sum(), (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_quarter_no).sum()]
    ]
    ACM_enforcement_table = pd.DataFrame(ACM_enforcement_table_data, columns=['Occupied buildings that have:', 'Undergone enforcement action', 'Undergone enforcement action supported by Joint Inspection Team', 'Forecast start available', ACM_enforcement_end_quarter_line])

    # Formatting the dataframe for the ACM year identified table
    Total_row = pd.DataFrame({
        'Year of identification': ['Total'],
        'Number of buildings identified': [ACM_11.loc[7, 'Current Month']],
        'Cumulative number': [ACM_11.loc[7, 'Current Month']],
    })
    ACM_yearly_identified = pd.concat([ACM_yearly_identified, Total_row], ignore_index=True)

    # Dictionary for transporting dataframes to the next script
    ACM_tables = {
        'ACM_remediation_table': ACM_remediation_table,
        'ACM_enforcement_table': ACM_enforcement_table,
        'ACM_yearly_identified': ACM_yearly_identified,
        'ACM_yearly_identified_line': ACM_yearly_identified_line
    }

    # Dictionary for transporting dataframes to the next script
    ACM_headline_dict = {
        'ACM_total': ACM_11.loc[7, 'Current Month'],
        'ACM_started_c_no': ACM_11.loc[3, 'Cumulative'],
        'ACM_started_c_pct': ACM_2.loc[3, 'Cumulative Percentage'],
        'ACM_started_line': Change_line_in_DR(ACM_11.loc[3, 'Cumulative Change']),
        'ACM_signoff_c_no': ACM_11.loc[1, 'Cumulative'],
        'ACM_signoff_c_pct': ACM_2.loc[1, 'Cumulative Percentage'],
        'ACM_signoff_line': Change_line_in_DR(ACM_11.loc[1, 'Cumulative Change']),
        'ACM_enforcement_total': convert_number((ACM_start_trajectory_filled_NaN['Vacant'] != 'N').sum() + (ACM_start_trajectory_filled_NaN['Vacant'] == 'N').sum()),
        'ACM_enforcement_pct': format_percentage(ACM_2.loc[4, 'Current Percentage'] + ACM_2.loc[5, 'Current Percentage'] + ACM_2.loc[6, 'Current Percentage']),
        'ACM_enforcement_line': Change_line_in_DR(ACM_11.loc[4, 'Change'] +  ACM_11.loc[5, 'Change'] + ACM_11.loc[6, 'Change']),
        'ACM_enforcement_vacants': convert_number((ACM_start_trajectory_filled_NaN['Vacant'] != 'N').sum()),
        'ACM_enforcement_forecast': convert_number((ACM_start_trajectory_nvt['StartedDatePlanned'] > datetime(2000, 1, 1)).sum()),
        'ACM_enforcement_enforced_no_forecast_word': convert_number(ACM_start_trajectory_nvt[(ACM_start_trajectory_nvt['Enforcement'] == 'Yes') & (ACM_start_trajectory_nvt['StartedDatePlanned'].isna())].shape[0]),
        'ACM_enforcement_not_enforced_no_forecast_word': convert_number(ACM_start_trajectory_nvt[(ACM_start_trajectory_nvt['Enforcement'] == 'No') & (ACM_start_trajectory_nvt['StartedDatePlanned'].isna())].shape[0])
    }

    # Dictitionary for transporting dataframes to the next script
    ACM_section_dict = {
        'ACM_complete_c_no': ACM_11.loc[0, 'Cumulative'],
        'ACM_signoff_c_no': ACM_11.loc[1, 'Cumulative'],
        'ACM_removed_c_no': ACM_11.loc[2, 'Cumulative'],
        'ACM_started_c_no': ACM_11.loc[3, 'Cumulative'],
        'ACM_complete_c_pct': ACM_2.loc[0, 'Cumulative Percentage'],
        'ACM_signoff_c_pct': ACM_2.loc[1, 'Cumulative Percentage'],
        'ACM_removed_c_pct': ACM_2.loc[2, 'Cumulative Percentage'],
        'ACM_started_c_pct': ACM_2.loc[3, 'Cumulative Percentage'],
        'ACM_total': ACM_11.loc[7, 'Current Month'],
        'ACM_complete_line': Change_line_in_DR(ACM_11.loc[0, 'Cumulative Change']),
        'ACM_signoff_line': Change_line_in_DR(ACM_11.loc[1, 'Cumulative Change']),
        'ACM_removed_line': Change_line_in_DR(ACM_11.loc[2, 'Cumulative Change']),
        'ACM_started_line': Change_line_in_DR(ACM_11.loc[3, 'Cumulative Change']),
        'ACM_total_line': Change_line_in_DR(ACM_11.loc[7, 'Change']),
        'ACM_enforcement_total': (
            (ACM_start_trajectory_filled_NaN['Vacant'] != 'N').sum() + 
            (ACM_start_trajectory_filled_NaN['Vacant'] == 'N').sum()
        ),
        'ACM_enforcement_pct': format_percentage(
            ACM_2.loc[4, 'Current Percentage'] + 
            ACM_2.loc[5, 'Current Percentage'] + 
            ACM_2.loc[6, 'Current Percentage']
        ),
        'ACM_enforcement_line': Change_line_in_DR(
            ACM_11.loc[4, 'Change'] + 
            ACM_11.loc[5, 'Change'] + 
            ACM_11.loc[6, 'Change']
        ),
        'ACM_enforcement_total_nvt': (
            (ACM_start_trajectory_filled_NaN['Vacant'] == 'N').sum()
        ),
        'ACM_enforcement_forecast_end_quarter_word': (
            convert_number(
                (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_quarter_no).sum()
            )
        ),
        'ACM_enforcement_forecast_this_year_word': (
            convert_number(
                (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_this_year).sum() - 
                (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_quarter_no).sum()
            )
        ),
        'ACM_enforcement_forecast_next_year_word': (
            convert_number(
                (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_next_year).sum() - 
                (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_this_year).sum()
            )
        ),
        'ACM_enforcement_enforced_end_quarter_word': (
            convert_number(
                ACM_start_trajectory_nvt[
                    (ACM_start_trajectory_nvt['Enforcement'] == 'Yes') & 
                    (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_quarter_no)
                ].shape[0]
            )
        ),
        'ACM_enforcement_enforced_this_year_word': (
            convert_number(
                ACM_start_trajectory_nvt[
                    (ACM_start_trajectory_nvt['Enforcement'] == 'Yes') & 
                    (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_this_year)
                ].shape[0] - ACM_start_trajectory_nvt[
                    (ACM_start_trajectory_nvt['Enforcement'] == 'Yes') & 
                    (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_quarter_no)
                ].shape[0]
            )
        ),
        'ACM_enforcement_enforced_next_year_word': (
            convert_number(
                ACM_start_trajectory_nvt[
                    (ACM_start_trajectory_nvt['Enforcement'] == 'Yes') & 
                    (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_next_year)
                ].shape[0] - ACM_start_trajectory_nvt[
                    (ACM_start_trajectory_nvt['Enforcement'] == 'Yes') & 
                    (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_this_year)
                ].shape[0]
            )
        ),
        'ACM_enforcement_enforced_no_forecast_word': (
            convert_number(
                ACM_start_trajectory_nvt[
                    (ACM_start_trajectory_nvt['Enforcement'] == 'Yes') & 
                    (ACM_start_trajectory_nvt['StartedDatePlanned'].isna())
                ].shape[0]
            )
        ),
        'ACM_enforcement_not_enforced_no_forecast_word': (
            convert_number(
                ACM_start_trajectory_nvt[
                    (ACM_start_trajectory_nvt['Enforcement'] == 'No') & 
                    (ACM_start_trajectory_nvt['StartedDatePlanned'].isna())
                ].shape[0]
            )
        ),
        'ACM_quarter_progress': format_percentage((ACM_11.loc[3, 'Cumulative'] + (ACM_start_trajectory_nvt['StartedDatePlanned'] < end_quarter_no).sum()) / ACM_11.loc[7, 'Current Month']),
        'ACM_social_started_c_pct': ACM_2.loc[3, 'Cumulative Social Percentage'],
        'ACM_private_started_c_pct': ACM_2.loc[3, 'Cumulative Private Percentage'],
        'ACM_social_total': ACM_2.loc[7, 'Social Number'],
        'ACM_private_total': ACM_2.loc[7, 'Private Number'],
        'ACM_current_yearly_identified': ACM_yearly_identified.iloc[-1]['Number of buildings identified'],
        'ACM_yearly_identified_total': ACM_yearly_identified_line['Number of buildings identified'].sum(),
        'ACM_completed_dwellings_low': format(ACM_7.loc[0, 'Low'], ','),
        'ACM_completed_dwellings_high': format(ACM_7.loc[0, 'High'], ','),
        'ACM_yet_to_dwellings_low': format(ACM_7.loc[4, 'Low'], ','),
        'ACM_yet_to_dwellings_high': format(ACM_7.loc[4, 'High'], ',')
    }

    return ACM_tables, ACM_headline_dict, ACM_section_dict
