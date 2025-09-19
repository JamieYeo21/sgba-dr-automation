# CSS_variables.py
"""
Created on Friday 17 January 2025, 14:32:06

Author: Harry Simmons
"""

from datetime import datetime
import calendar as cal
import pandas as pd
import sys
import os
import inflect

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
# Navigate to the Utility folder relative to the script
utility_path = os.path.join(script_dir, 'Utility')
# Add it to sys.path so that python can import from it
sys.path.append(utility_path)


from CSS.CSS_data_handler import CSS_retrieve_data
from Utility.functions import convert_number, format_percentage, Change_line_in_DR, number_or_none

def CSS_variable_creator (CSS_handled_data):
    CSS_1a = CSS_handled_data['CSS_1a']
    CSS_1b = CSS_handled_data['CSS_1b']
    CSS_2 = CSS_handled_data['CSS_2']
    CSS_3 = CSS_handled_data['CSS_3']
    CSS_4 = CSS_handled_data['CSS_4']
    CSS_5_current  = CSS_handled_data['CSS_5_current']

    CSS_5_previous = CSS_handled_data['CSS_5_previous']
    CSS_6 = CSS_handled_data['CSS_6']
    CSS_misc = CSS_handled_data['CSS_misc']

    CSS_remediation_table = pd.DataFrame({
        'Remediation Stage': ['Remediation complete', 'Works started', 'In programme', 'Total'],
        'Number of buildings': CSS_6['Current Month'],
        'Percentage': CSS_1b['Current Percentage'],
        'Cumulative Number': CSS_6['Cumulative'],
        'Cumulative Percentage': CSS_1b['Cumulative Percentage']
    })
    
    CSS_tables = {
        'CSS_remediation_table' : CSS_remediation_table
    }

    CSS_headline_dict = {
        'CSS_eligible_total': format(CSS_6.loc[3, 'Current Month'], ','),
        'CSS_BSF_transfer': format(CSS_misc.loc[1, 'Current Month'], ','),
        'CSS_eligible_total_line': Change_line_in_DR(CSS_6.loc[3, 'Change']),
        'CSS_started_c_no': format(CSS_6.loc[1, 'Cumulative'], ','),
        'CSS_started_c_pct': CSS_1b.loc[1, 'Cumulative Percentage'],
        'CSS_started_c_line': Change_line_in_DR(CSS_6.loc[1, 'Cumulative Change']),
        'CSS_completed_nc_no': format(CSS_6.loc[0, 'Cumulative'], ','),
        'CSS_completed_c_pct': CSS_1b.loc[0, 'Cumulative Percentage'],
        'CSS_completed_c_line': Change_line_in_DR(CSS_6.loc[0, 'Change']),
        'CSS_pre_eligible_total': format(CSS_1a.loc[2, 'Number'], ','),
        'CSS_pre_eligible': format(CSS_1a.loc[1, 'Number'], ','),
        'CSS_pre_application': format(CSS_1a.loc[0, 'Number'], ',')
    }

    CSS_section_dict = {
        'CSS_total_total': format(CSS_1a.iloc[2, 1] + CSS_6.loc[3, 'Current Month'], ','),
        'CSS_pre_eligible': format(CSS_1a.iloc[1, 1], ','),
        'CSS_eligible_total': format(CSS_6.loc[3, 'Current Month'], ','),
        'CSS_started_c_no': format(CSS_6.loc[1, 'Cumulative'], ','),
        'CSS_started_c_pct': CSS_1b.loc[1, 'Cumulative Percentage'],
        'CSS_completed_c_no': format(CSS_6.loc[0, 'Cumulative'], ','),
        'CSS_completed_c_pct': CSS_1b.loc[0, 'Cumulative Percentage'],
        'CSS_pre_eligible_total': format(CSS_1a.iloc[2, 1], ','),
        'CSS_pre_application': format(CSS_1a.iloc[0, 1], ','),
        'CSS_eligible_total_line': Change_line_in_DR(CSS_6.loc[3, 'Change']),
        'CSS_BSF_transfer': format(CSS_misc.loc[1, 'Current Month'], ','),
        'CSS_BSF_transfer_line': number_or_none(CSS_misc.loc[1, 'Change']),
        'CSS_GFA': format(CSS_misc.loc[0, 'Current Month'], ','),
        'CSS_GFA_pct': format_percentage(CSS_misc.loc[0, 'Current Month'] / CSS_6.loc[3, 'Current Month']),
        'CSS_GFA_line': Change_line_in_DR(CSS_misc.loc[0, 'Change']),

        'CSS_PTSP': format(CSS_5_current.loc[0, 'Total'], ','),
        'CSS_PTSP_pct': format_percentage(CSS_5_current.loc[0, 'Total'] / CSS_6.loc[3, 'Current Month']),
        'CSS_PTSP_line': Change_line_in_DR((CSS_5_current.loc[0, 'Total'] - CSS_5_previous.loc[0, 'Total'])),
        
        'CSS_started_c_line': Change_line_in_DR(CSS_6.loc[1, 'Cumulative Change']),
        'CSS_completed_c_line': Change_line_in_DR(CSS_6.loc[0, 'Change']),

        'CSS_northern_ireland': number_or_none(CSS_2.iloc[9, 1]),
        'CSS_18m_c_pct': format_percentage(CSS_3.loc[1, '18m Cumulative Percentage']),
        'CSS_11_18m_c_pct': format_percentage(CSS_3.loc[1, '11_18m Cumulative Percentage']),
        'CSS_private_c_pct': format_percentage(CSS_4.loc[1, 'Private Cumulative Percentage']),
        'CSS_social_c_pct': format_percentage(CSS_4.loc[1, 'Social Cumulative Percentage']),
    }

    return CSS_tables, CSS_headline_dict, CSS_section_dict