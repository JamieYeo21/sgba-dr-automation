# RAP_variables.py
"""
Created on Tuesday 26 August 2025, 11:18:21

Author: Matthew Bandura
"""

import pandas as pd
import os
import sys

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
# Navigate to the Utility folder relative to the script
utility_path = os.path.join(script_dir, 'Utility')
# Add it to sys.path so that python can import from it
sys.path.append(utility_path)


from RAP.RAP_data_handler import RAP_retrieve_data
from Utility.functions import format_percentage

def RAP_variable_creator(RAP_handled_data):
    # Unpacking dataframes from RAP_data_handler
    Combined_2 = RAP_handled_data['Combined_2']
    Estimated_2 = RAP_handled_data['Estimated_2']
    Estimated_4 = RAP_handled_data['Estimated_4']
    Combined_8 = RAP_handled_data['Combined_8']
    RAP_misc = RAP_handled_data['RAP_misc']


    RAP_section_dict = {
        'RAP_18m_complete_no': format(Combined_8.loc[0, '18m Number'], ','),
        'RAP_18m_underway_no': format(Combined_8.loc[1, '18m Number'], ','),
        'RAP_18m_programme_no': format(Combined_8.loc[2, '18m Number'],','),

        'RAP_18m_est_complete_high_pct' : format_percentage(Combined_8.loc[0, '18m Number'] / RAP_misc.loc[0, 'Number']).replace('%', ''),
        'RAP_18m_est_complete_low_pct' : format_percentage(Combined_8.loc[0, '18m Number'] / RAP_misc.loc[1, 'Number']),

        'RAP_18m_est_underway_high_pct' : format_percentage(Combined_8.loc[1, '18m Number'] / RAP_misc.loc[0, 'Number']).replace('%', ''),
        'RAP_18m_est_underway_low_pct' : format_percentage(Combined_8.loc[1, '18m Number'] / RAP_misc.loc[1, 'Number']),

        'RAP_11m_total_complete_no' : format(Combined_2.loc[0, 'Current Number'], ','),
        'RAP_11m_total_programme_no' : format((Combined_2.loc[1, 'Current Number'] + Combined_2.loc[2, 'Current Number']), ','),

        'RAP_11m_est_complete_high_pct' : format_percentage((Combined_2.loc[0, 'Current Number']) / Estimated_2.loc[2, 'High Estimate']).replace('%', ''),
        'RAP_11m_est_complete_low_pct'  : format_percentage((Combined_2.loc[0, 'Current Number']) / Estimated_2.loc[2, 'Low Estimate']),

        'RAP_11m_est_programme_high_pct' : format_percentage((Combined_2.loc[1, 'Current Number'] + Combined_2.loc[2, 'Current Number']) / Estimated_2.loc[2, 'High Estimate']).replace('%', ''),
        'RAP_11m_est_programme_low_pct' : format_percentage((Combined_2.loc[1, 'Current Number'] + Combined_2.loc[2, 'Current Number']) / Estimated_2.loc[2, 'Low Estimate']),

        'RAP_11m_est_remaining_low_no' : format((Estimated_2.loc[2, 'Low Estimate'] - Combined_2.loc[3,'Current Number'] ), ','),
        'RAP_11m_est_remaining_high_no' : format((Estimated_2.loc[2, 'High Estimate'] - Combined_2.loc[3,'Current Number']), ','),

        'RAP_11m_est_remaining_low_pct' : format_percentage((Estimated_2.loc[2, 'Low Estimate'] - Combined_2.loc[3,'Current Number']) / Estimated_2.loc[2, 'Low Estimate']).replace('%', ''),
        'RAP_11m_est_remaining_high_pct' : format_percentage((Estimated_2.loc[2, 'High Estimate'] - Combined_2.loc[3,'Current Number']) / Estimated_2.loc[2, 'High Estimate']),
    }

    return RAP_section_dict
