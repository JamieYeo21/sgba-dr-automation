# Developer_variables.py
"""
Created on Wednesday 19 February 2025, 14:39:58

Author: Harry Simmons
"""

from datetime import datetime
import calendar as cal
import pandas as pd
import sys
import os

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
# Navigate to the Utility folder relative to the script
utility_path = os.path.join(script_dir, 'Utility')
# Add it to sys.path so that python can import from it
sys.path.append(utility_path)


from Developer.Developer_data_handler import Developer_retrieve_data_last_month, Developer_retrieve_data_this_month
from Utility.functions import Change_line_in_DR, format_percentage, convert_number
from Utility.dates import sort_dates

def Developer_variable_creator(Developer_handled_data_last_month, Developer_handled_data_this_month, dates_variables):
    # Unpack dates variables
    month = dates_variables['month']
    
    # Unpack last months dfs
    Developer_1a_last_month = Developer_handled_data_last_month['Developer_1a']
    Developer_1b_last_month = Developer_handled_data_last_month['Developer_1b']
    Developer_1c_last_month = Developer_handled_data_last_month['Developer_1c']
    Developer_1d_last_month = Developer_handled_data_last_month['Developer_1d']
    Developer_2_last_month = Developer_handled_data_last_month['Developer_2']

    # Unpack this months dfs
    Developer_1a_this_month = Developer_handled_data_this_month['Developer_1a']
    Developer_1b_this_month = Developer_handled_data_this_month['Developer_1b']
    Developer_1c_this_month = Developer_handled_data_this_month['Developer_1c']
    Developer_1d_this_month = Developer_handled_data_this_month['Developer_1d']
    Developer_2_this_month = Developer_handled_data_this_month['Developer_2']
    Developer_3_this_month = Developer_handled_data_this_month['Developer_3']
    Developer_misc = Developer_handled_data_this_month['Developer_misc']

    # Making the data changed line
    if month in [11, 8, 5, 2]:
        Developer_data_change_line = ''
    else:
        Developer_data_change_line = ' and remains unchanged from the previous publication'

    #unpacking the variables from sort_dates
    Developer_last_day = dates_variables['dev_last_day']
    Developer_month = dates_variables['dev_month']
    Developer_year = dates_variables['dev_year']

    # Work out how many developers
    other_row = Developer_3_this_month[Developer_3_this_month.iloc[:, 0].str.contains('other', case=False, na=False)].index.tolist()
    other_row = other_row[0]
    Developer_number_of = other_row - 1

    # Formatting the dataframe for Table 7
    Developer_remediation_table = pd.DataFrame({
        'Remediation Stage': ['Remediation complete', 'Remediation complete – awaiting building control sign-off', 'Remediation started', 'Remediation not started – plans in place', 'Remediation not started – no plans in place', 'Total'],
        'Number of buildings': Developer_1a_this_month['Current Number'],
        'Percentage': Developer_1a_this_month['Current'],
        'Cumulative Number': Developer_1a_this_month['Cumulative Number'],
        'Cumulative Percentage': Developer_1a_this_month['Cumulative']
    })

    Developer_tables = {
        'Developer_remediation_table' : Developer_remediation_table
    }
    
    Developer_headline_dict = {
        'Developer_cutoff': dates_variables['dev_cutoff'],

        'Developer_life_critical_total': format(Developer_1a_this_month.iloc[5, 5], ','),
        'Developer_life_critical_total_line': Change_line_in_DR(Developer_1a_this_month.iloc[5, 5] - Developer_1a_last_month.iloc[5, 5]),

        'Developer_started_c_no': format(Developer_1a_this_month.loc[2, 'Cumulative Number'], ','),
        'Developer_started_c_pct': Developer_1a_this_month.loc[2, 'Cumulative'],
        'Developer_started_line': Change_line_in_DR(Developer_1a_this_month.loc[2, 'Cumulative Number'] - Developer_1a_last_month.loc[2, 'Cumulative Number']),

        'Developer_signoff_c_no': format(Developer_1a_this_month.loc[1, 'Cumulative Number'], ','),
        'Developer_signoff_c_pct': Developer_1a_this_month.loc[1, 'Cumulative'],
        'Developer_signoff_line': Change_line_in_DR(Developer_1a_this_month.loc[1, 'Cumulative Number'] - Developer_1a_last_month.loc[1, 'Cumulative Number']),

        'Developer_cladding_defects_total': format(Developer_1b_this_month.iloc[5, 5], ','),
        'Developer_cladding_defects_total_line': Change_line_in_DR(Developer_1b_this_month.iloc[5, 5] - Developer_1b_last_month.iloc[5, 5]),

        'Developer_cladding_defects_started_c_no': format(Developer_1b_this_month.loc[2, 'Cumulative Number'], ','),
        'Developer_cladding_defects_started_c_pct': Developer_1b_this_month.loc[2, 'Cumulative'],
        'Developer_cladding_defects_started_line': Change_line_in_DR(Developer_1b_this_month.loc[2, 'Cumulative Number'] - Developer_1b_last_month.loc[2, 'Cumulative Number']),

        'Developer_cladding_defects_signoff_c_no': format(Developer_1b_this_month.loc[1, 'Cumulative Number'], ','),
        'Developer_cladding_defects_signoff_c_pct': Developer_1b_this_month.loc[1, 'Cumulative'],
        'Developer_cladding_defects_signoff_line': Change_line_in_DR(Developer_1b_this_month.loc[1, 'Cumulative Number'] - Developer_1b_last_month.loc[1, 'Cumulative Number']),
    }

    Developer_section_dict = {
        'Developer_cutoff': dates_variables['dev_cutoff'],

        'Developer_started_c_pct': Developer_1a_this_month.loc[2, 'Cumulative'],
        'Developer_signoff_c_pct': Developer_1a_this_month.loc[1, 'Cumulative'],

        'Developer_data_change_line': Developer_data_change_line,

        'Developer_number_of': Developer_number_of,

        'Developer_responsibility_total': format(Developer_3_this_month.iloc[0, 1], ','),
        
        'Developer_life_critical_total': format(Developer_1a_this_month.iloc[5, 5], ','),
        'Developer_life_critical_total_line': Change_line_in_DR(Developer_1a_this_month.iloc[5, 5] - Developer_1a_last_month.iloc[5, 5]),

        'Developer_signoff_c_no': format(Developer_1a_this_month.loc[1, 'Cumulative Number'], ','),
        'Developer_signoff_line': Change_line_in_DR(Developer_1a_this_month.loc[1, 'Cumulative Number'] - Developer_1a_last_month.loc[1, 'Cumulative Number']),

        'Developer_complete_c_no': format(Developer_1a_this_month.loc[0, 'Cumulative Number'], ','),
        'Developer_complete_c_pct': Developer_1a_this_month.loc[0, 'Cumulative'],

        'Developer_started_c_no': format(Developer_1a_this_month.loc[2, 'Cumulative Number'], ','),
        'Developer_started_line': Change_line_in_DR(Developer_1a_this_month.loc[2, 'Cumulative Number'] - Developer_1a_last_month.loc[2, 'Cumulative Number']),
    
        'Developer_plans_c_no': format(Developer_1a_this_month.loc[3, 'Current Number'], ','),
        'Developer_plans_c_pct': Developer_1a_this_month.loc[3, 'Current'],
        'Developer_plans_line': Change_line_in_DR(Developer_1a_this_month.loc[3, 'Current Number'] - Developer_1a_last_month.loc[3, 'Current Number']),

        'Developer_no_plans_c_no': format(Developer_1a_this_month.loc[4, 'Current Number'], ','),
        'Developer_no_plans_c_pct': Developer_1a_this_month.loc[4, 'Current'],
        'Developer_no_plans_line': Change_line_in_DR(Developer_1a_this_month.loc[4, 'Current Number'] - Developer_1a_last_month.loc[4, 'Current Number']),

        'Developer_cost': round(((Developer_2_this_month.iloc[0, 3]) / 1000), 1),
        'Developer_cost_change': round(((Developer_2_this_month.iloc[0, 3] - Developer_2_last_month.iloc[0, 3]) / 1000), 1),

        'Developer_dwellings_total': format(round(Developer_misc.iloc[0, 1]), ','),
        'Developer_dwellings_started_c_no': format(Developer_misc.iloc[1, 1], ','),

        'Developer_BSF_transfer_total': format(Developer_1d_this_month.loc[6, 'Current Number'], ','),
        'Developer_BSF_transfer_total_line': Change_line_in_DR(Developer_1d_this_month.loc[6, 'Current Number'] - Developer_1d_last_month.loc[6, 'Current Number']),

        'Developer_BSF_transfer_complete_c_no': format(Developer_1d_this_month.loc[0, 'Cumulative Number'], ','),
        'Developer_BSF_transfer_complete_c_pct': Developer_1d_this_month.loc[0, 'Cumulative Percentage'],

        'Developer_BSF_transfer_signoff_c_no': format(Developer_1d_this_month.loc[1, 'Cumulative Number'], ','),
        'Developer_BSF_transfer_signoff_c_pct': Developer_1d_this_month.loc[1, 'Cumulative Percentage'],
        'Developer_BSF_transfer_signoff_line': Change_line_in_DR(Developer_1d_this_month.loc[1, 'Cumulative Number'] - Developer_1d_last_month.loc[1, 'Cumulative Number']),

        'Developer_BSF_transfer_started_c_no': format(Developer_1d_this_month.loc[2, 'Cumulative Number'], ','),
        'Developer_BSF_transfer_started_c_pct': Developer_1d_this_month.loc[2, 'Cumulative Percentage'],
        'Developer_BSF_transfer_started_line': Change_line_in_DR(Developer_1d_this_month.loc[2, 'Cumulative Number'] - Developer_1d_last_month.loc[2, 'Cumulative Number']),

        'Developer_BSF_transfer_plans_nc_no': format(Developer_1d_this_month.loc[3, 'Current Number'], ','),
        'Developer_BSF_transfer_plans_nc_pct': Developer_1d_this_month.loc[3, 'Current Percentage'],
        'Developer_BSF_transfer_plans_line': Change_line_in_DR(Developer_1d_this_month.loc[3, 'Current Number'] - Developer_1d_last_month.loc[3, 'Current Number']),

        'Developer_BSF_transfer_no_plans_nc_no': format(Developer_1d_this_month.loc[4, 'Current Number'], ','),
        'Developer_BSF_transfer_no_plans_nc_pct': Developer_1d_this_month.loc[4, 'Current Percentage'],
        'Developer_BSF_transfer_no_plans_line': Change_line_in_DR(Developer_1d_this_month.loc[4, 'Current Number'] - Developer_1d_last_month.loc[4, 'Current Number']),

        'Developer_BSF_transfer_not_life_critical_nc_no': format(Developer_1d_this_month.loc[5, 'Current Number'], ','),
        'Developer_BSF_transfer_not_life_critical_nc_pct': Developer_1d_this_month.loc[5, 'Current Percentage'],
        'Developer_BSF_transfer_not_life_critical_line': Change_line_in_DR(Developer_1d_this_month.loc[5, 'Current Number'] - Developer_1d_last_month.loc[5, 'Current Number']),

        'Developer_forecast_timeframe': f"1 {cal.month_name[Developer_month + 1]} {Developer_year} and {Developer_last_day} {cal.month_name[Developer_month]} {Developer_year + 1}",
        'Developer_forecast_starts': format(Developer_misc.iloc[2, 1], ','),
        'Developer_forecast_completes': format(Developer_misc.iloc[3, 1], ','),

        'Developer_11_18m_started_c_pct': Developer_1a_this_month.loc[2, 'Cumulative 11_18m Percentage'],
        'Developer_18m_started_c_pct': Developer_1a_this_month.loc[2, 'Cumulative 18m Percentage'],

        'Developer_cladding_defects_total': format(Developer_1b_this_month.iloc[5, 5], ','),
        'Developer_cladding_defects_total_line': Change_line_in_DR(Developer_1b_this_month.iloc[5, 5] - Developer_1b_last_month.iloc[5, 5]),

        'Developer_cladding_defects_signoff_c_no': format(Developer_1b_this_month.loc[1, 'Cumulative Number'], ','),
        'Developer_cladding_defects_signoff_c_pct': Developer_1b_this_month.loc[1, 'Cumulative'],
        'Developer_cladding_defects_signoff_line': Change_line_in_DR(Developer_1b_this_month.loc[1, 'Cumulative Number'] - Developer_1b_last_month.loc[1, 'Cumulative Number']),

        'Developer_cladding_defects_complete_c_no': format(Developer_1b_this_month.loc[0, 'Cumulative Number'], ','),
        'Developer_cladding_defects_complete_c_pct': Developer_1b_this_month.loc[0, 'Cumulative'],

        'Developer_cladding_defects_started_c_no': format(Developer_1b_this_month.loc[2, 'Cumulative Number'], ','),
        'Developer_cladding_defects_started_c_pct': Developer_1b_this_month.loc[2, 'Cumulative'],
        'Developer_cladding_defects_started_line': Change_line_in_DR(Developer_1b_this_month.loc[2, 'Cumulative Number'] - Developer_1b_last_month.loc[2, 'Cumulative Number']),

        'Developer_cladding_defects_plans_c_no': format(Developer_1b_this_month.loc[3, 'Current Number'], ','),
        'Developer_cladding_defects_plans_c_pct': Developer_1b_this_month.loc[3, 'Current'],
        'Developer_cladding_defects_plans_line': Change_line_in_DR(Developer_1b_this_month.loc[3, 'Current Number'] - Developer_1b_last_month.loc[3, 'Current Number']),

        'Developer_cladding_defects_no_plans_c_no': format(Developer_1b_this_month.loc[4, 'Current Number'], ','),
        'Developer_cladding_defects_no_plans_c_pct': Developer_1b_this_month.loc[4, 'Current'],
        'Developer_cladding_defects_no_plans_line': Change_line_in_DR(Developer_1b_this_month.loc[4, 'Current Number'] - Developer_1b_last_month.loc[4, 'Current Number']),

        'Developer_self_reported_total': format(Developer_1c_this_month.iloc[5, 5], ','),
        'Developer_self_reported_total_line': Change_line_in_DR(Developer_1c_this_month.iloc[5, 5] - Developer_1c_last_month.iloc[5, 5]),

        'Developer_self_reported_signoff_c_no': format(Developer_1c_this_month.loc[1, 'Cumulative Number'], ','),
        'Developer_self_reported_signoff_c_pct': Developer_1c_this_month.loc[1, 'Cumulative'],

        'Developer_self_reported_complete_c_no': format(Developer_1c_this_month.loc[0, 'Cumulative Number'], ','),
        'Developer_self_reported_complete_c_pct': Developer_1c_this_month.loc[0, 'Cumulative'],

        'Developer_self_reported_signoff_line': Change_line_in_DR(Developer_1c_this_month.loc[1, 'Cumulative Number'] - Developer_1c_last_month.loc[1, 'Cumulative Number']),

        'Developer_self_reported_started_c_no': format(Developer_1c_this_month.loc[2, 'Cumulative Number'], ','),
        'Developer_self_reported_started_c_pct': Developer_1c_this_month.loc[2, 'Cumulative'],
        'Developer_self_reported_started_line': Change_line_in_DR(Developer_1c_this_month.loc[2, 'Cumulative Number'] - Developer_1c_last_month.loc[2, 'Cumulative Number']),

        'Developer_self_reported_plans_c_no': format(Developer_1c_this_month.loc[3, 'Current Number'], ','),
        'Developer_self_reported_plans_c_pct': Developer_1c_this_month.loc[3, 'Current'],
        'Developer_self_reported_plans_line': Change_line_in_DR(Developer_1c_this_month.loc[3, 'Current Number'] - Developer_1c_last_month.loc[3, 'Current Number']),

        'Developer_self_reported_no_plans_c_no': format(Developer_1c_this_month.loc[4, 'Current Number'], ','),
        'Developer_self_reported_no_plans_c_pct': Developer_1c_this_month.loc[4, 'Current'],
        'Developer_self_reported_no_plans_line': Change_line_in_DR(Developer_1c_this_month.loc[4, 'Current Number'] - Developer_1c_last_month.loc[4, 'Current Number']),
   }

    return Developer_tables, Developer_headline_dict, Developer_section_dict


