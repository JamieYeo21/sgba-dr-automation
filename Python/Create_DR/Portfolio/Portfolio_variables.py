# Portfolio_variables.py
"""
Created on Monday 27 January 2025, 09:14:43

Author: Harry Simmons
"""

import pandas as pd
import os
import sys

# Add the Utility folder to sys.path
folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))
sys.path.append(folder_path)

from Portfolio.Portfolio_data_handler import Portfolio_retrieve_data
from Utility.functions import format_percentage, convert_number, more_or_fewer, Change_line_in_DR
from Utility.dates import sort_dates

def Portfolio_variable_creator(Portfolio_handled_data):
    # Unpacking dataframes from ACM_data_handler
    Combined_2 = Portfolio_handled_data['Combined_2']
    Combined_4 = Portfolio_handled_data['Combined_4']
    Combined_5 = Portfolio_handled_data['Combined_5']
    Combined_6 = Portfolio_handled_data['Combined_6']

    Portfolio_remediation_table = pd.DataFrame({
        'Remediation Stage': ['Remediation complete', 'Remediation underway', 'In programme', 'Total'],
        'Number of buildings': Combined_6['Current Month'],
        'Percentage': Combined_2['Current Percentage'],
        'Cumulative Number': Combined_6['Cumulative'],
        'Cumulative Percentage': Combined_2['Cumulative Percentage']
    })
    Portfolio_tables = {
        'Portfolio_remediation_table' : Portfolio_remediation_table
    }
    Portfolio_headline_dict = {
        'Portfolio_total': format(Combined_6.loc[3, 'Current Month'], ','),
        'Portfolio_total_line': Change_line_in_DR(Combined_6.loc[3, 'Monthly Change']),
        'Portfolio_total_since_oct_23': more_or_fewer(Combined_6.loc[3, 'Since October 2023']),
        'Portfolio_started_c_no': format(Combined_6.loc[1, 'Cumulative'], ','),
        'Portfolio_started_c_pct': Combined_2.loc[1, 'Cumulative Percentage'],
        'Portfolio_completed_c_no': format(Combined_6.loc[0, 'Current Month'], ','),
        'Portfolio_completed_c_pct': Combined_2.loc[0, 'Cumulative Percentage'],
    }

    Portfolio_section_dict = {
        'Portfolio_started_c_no': format(Combined_6.loc[1, 'Cumulative'], ','),
        'Portfolio_started_c_pct': Combined_2.loc[1, 'Cumulative Percentage'],
        'Portfolio_completed_c_no': format(Combined_6.loc[0, 'Current Month'], ','),
        'Portfolio_completed_c_pct': Combined_2.loc[0, 'Cumulative Percentage'],
        'Portfolio_total': format(Combined_6.loc[3, 'Current Month'], ','),
        'Portfolio_started_nc_no': format(Combined_6.loc[1, 'Current Month'], ','),
        'Portfolio_started_nc_pct': Combined_2.loc[1, 'Current Percentage'],
        'Portfolio_in_programme_nc_no': format(Combined_6.loc[2, 'Current Month'], ','),
        'Portfolio_in_programme_nc_pct': Combined_2.loc[2, 'Current Percentage'],
        'Portfolio_total_monthly_change': more_or_fewer(Combined_6.loc[3, 'Monthly Change']),
        'Portfolio_started_c_monthly_change': more_or_fewer(Combined_6.loc[1, 'Cumulative Monthly Change']),
        'Portfolio_completed_c_monthly_change': more_or_fewer(Combined_6.loc[0, 'Cumulative Monthly Change']),
        'Portfolio_total_yearly_change': more_or_fewer(Combined_6.loc[3, 'Yearly Change']),
        'Portfolio_started_c_yearly_change': more_or_fewer(Combined_6.loc[1, 'Cumulative Yearly Change']),
        'Portfolio_completed_c_yearly_change': more_or_fewer(Combined_6.loc[0, 'Cumulative Yearly Change']),
        'Portfolio_total_dwellings': format(Combined_4.loc[3, 'Total Dwellings'], ','),
        'Portfolio_completed_dwellings': format(Combined_4.loc[0, 'Total Dwellings'], ','),
        'Portfolio_started_dwellings': format(Combined_4.loc[1, 'Total Dwellings'], ','),
        'Portfolio_in_programme_dwellings': format(Combined_4.loc[2, 'Total Dwellings'], ','),
        'Portfolio_11_18m_started_c_pct': Combined_2.loc[1, 'Cumulative 11_18m Percentage'],
        'Portfolio_18m_started_c_pct': Combined_2.loc[1, 'Cumulative 18m Percentage'],
        'Portfolio_social_started_c_pct': Combined_5.loc[1, 'Cumulative Social Percentage'],
        'Portfolio_private_started_c_pct': Combined_5.loc[1, 'Cumulative Private Percentage']
    }

    return Portfolio_tables, Portfolio_headline_dict, Portfolio_section_dict
