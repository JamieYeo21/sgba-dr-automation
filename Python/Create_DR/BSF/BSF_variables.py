# BSF_variables.py
"""
Created on Thursday 09 January 2025, 10:39:17

Author: Harry Simmons
"""

from datetime import datetime
import calendar as cal
import pandas as pd
import sys
import os
import inflect

# Add the Utility folder to sys.path
folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))
sys.path.append(folder_path)

from BSF.BSF_data_handler import BSF_retrieve_data
from Utility.functions import convert_number, format_percentage, Change_line_in_DR, more_or_fewer

def BSF_variable_creator (BSF_handled_data):
    #Unpacking the data
    BSF_1 = BSF_handled_data['BSF_1']
    BSF_5 = BSF_handled_data['BSF_5']
    BSF_reg_status = BSF_handled_data['BSF_reg_status']
    BSF_misc = BSF_handled_data['BSF_misc']

    BSF_remediation_table = pd.DataFrame({
        'Remediation Stage' : ['Remediation complete', 'Remediation complete: awaiting building control sign-off', 'Remediation started', 'Remediation plans in place', 'Intent to remediate', 'Total'],
        'Number of buildings' : BSF_5['Current Month'],
        'Percentage' : BSF_1['Current Percentage'],
        'Cumulative Number' : BSF_5['Cumulative'],
        'Cumulative Percentage' : BSF_1['Cumulative Percentage'] 
    })

    BSF_tables = {
        'BSF_remediation_table' : BSF_remediation_table
    }
    
    BSF_headline_dict = {
        'BSF_BSF_5_total': BSF_5.loc[5, 'Current Month'],
        'BSF_started_nc_no': BSF_5.loc[2, 'Current Month'],
        'BSF_started_nc_pct': BSF_1.loc[2, 'Current Percentage'],
        'BSF_signoff_c_no': BSF_5.loc[1, 'Cumulative'],
        'BSF_signoff_c_pct': BSF_1.loc[1, 'Cumulative Percentage'],
        'BSF_started_c_no': BSF_5.loc[2, 'Cumulative'],
        'BSF_started_c_pct': BSF_1.loc[2, 'Cumulative Percentage'],
        'BSF_started_c_line': Change_line_in_DR(BSF_5.loc[2, 'Cumulative Monthly Change']),
        'BSF_signoff_c_line': Change_line_in_DR(BSF_5.loc[1, 'Cumulative Monthly Change']),
    }

    BSF_section_dict = {
        'BSF_signoff_c_pct': BSF_1.loc[1, 'Cumulative Percentage'],
        'BSF_started_c_pct': BSF_1.loc[2, 'Cumulative Percentage'],
        'BSF_BSF_5_total': BSF_5.loc[5, 'Current Month'],
        'BSF_CSS_transfers_last_month': format(BSF_reg_status.loc[1, 'Last Month'], ','),
        'BSF_CSS_transfers_this_month': format(BSF_reg_status.loc[1, 'Current Month'], ','),
        'BSF_CSS_transfers_line': Change_line_in_DR(BSF_reg_status.loc[1, 'Current Month'] - BSF_reg_status.loc[1, 'Last Month']),
        'BSF_ineligible': format(BSF_reg_status.loc[2, 'Current Month'], ','),
        'BSF_withdrawn': format(BSF_reg_status.loc[3, 'Current Month'], ','),
        'BSF_developer_transfers': format(BSF_reg_status.loc[4, 'Current Month'], ','),
        'BSF_insufficient_evidence': format(BSF_reg_status.loc[5, 'Current Month'] + BSF_reg_status.loc[6, 'Current Month'], ','),
        'BSF_BSF_1_total': format(BSF_reg_status.loc[8, 'Current Month'] - BSF_reg_status.loc[0, 'Current Month'], ','),
        'BSF_developer_reimbursed_word': convert_number(BSF_misc.loc[0, 'Number']),
        'BSF_developer_anticipated_word': convert_number(BSF_misc.loc[1, 'Number']),
        'BSF_developer_reimbursed_pct': format_percentage(BSF_misc.loc[0, 'Number'] / BSF_5.loc[5, 'Current Month']),
        'BSF_developer_anticipated_pct': format_percentage(BSF_misc.loc[1, 'Number'] / BSF_5.loc[5, 'Current Month']),
        'BSF_FRAEW': BSF_misc.loc[3, 'Number'],
        'BSF_CAN': BSF_misc.loc[4, 'Number'],
        'BSF_started_c_no': BSF_5.loc[2, 'Cumulative'],
        'BSF_started_c_line': Change_line_in_DR(BSF_5.loc[2, 'Cumulative Monthly Change']),
        'BSF_started_nc_no': BSF_5.loc[2, 'Current Month'],
        'BSF_started_nc_pct': BSF_1.loc[2, 'Current Percentage'],
        'BSF_signoff_c_no': BSF_5.loc[1, 'Cumulative'],
        'BSF_signoff_c_line': Change_line_in_DR(BSF_5.loc[1, 'Cumulative Monthly Change']),
        'BSF_complete_nc_no': BSF_5.loc[0, 'Current Month'],
        'BSF_complete_nc_pct': BSF_1.loc[0, 'Current Percentage'],
        'BSF_not_yet_started': BSF_5.loc[3, 'Current Month'] + BSF_5.loc[4, 'Current Month'],
        'BSF_plans_nc_no': BSF_5.loc[3, 'Current Month'],
        'BSF_plans_nc_pct': BSF_1.loc[3, 'Current Percentage'],
        'BSF_intent_nc_no': BSF_5.loc[4, 'Current Month'],
        'BSF_intent_nc_pct': BSF_1.loc[4, 'Current Percentage'],
        'BSF_dwellings': format(round(BSF_misc.loc[2, 'Number'], -3), ','),
        'BSF_figure12_eligable_change': more_or_fewer(BSF_5.loc[5, 'Cumulative Yearly Change']),
        'BSF_figure12_started_c_change': more_or_fewer(BSF_5.loc[2, 'Cumulative Yearly Change']),
        'BSF_figure12_completed_c_change': more_or_fewer(BSF_5.loc[1, 'Cumulative Yearly Change']),
        'BSF_social_started_pct': BSF_1.loc[2, 'Cumulative Social Percentage'],
        'BSF_private_started_pct': BSF_1.loc[2, 'Cumulative Private Percentage'],
        'BSF_social_complete_pct': BSF_1.loc[1, 'Cumulative Social Percentage'],
        'BSF_private_complete_pct': BSF_1.loc[1, 'Cumulative Private Percentage']
    }

    BSF_developer_transfers =  format(BSF_reg_status.loc[4, 'Current Month'], ',')

    return BSF_tables, BSF_headline_dict, BSF_section_dict, BSF_developer_transfers