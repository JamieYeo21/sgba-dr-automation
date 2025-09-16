# Enforcement_variables.py
"""
Created on Tuesday 04 March 2025, 08:43:31

Author: Harry Simmons
"""

from datetime import datetime
import calendar as cal
import pandas as pd
import sys
import os
import re

# Add the Utility folder to sys.path
folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))
sys.path.append(folder_path)

from Enforcement.Enforcement_data_handler import Enforcement_retrieve_data
from Utility.functions import convert_number, format_percentage, Change_line_in_DR, number_or_none

def Enforcement_variable_creator(Enforcement_handled_data):
    # Unpack df's
    Enforcement_1_uncut = Enforcement_handled_data['Enforcement_1_uncut']
    Enforcement_1 = Enforcement_handled_data['Enforcement_1']
    Enforcement_2 = Enforcement_handled_data['Enforcement_2']
    Enforcement_3 = Enforcement_handled_data['Enforcement_3']
    Enforcement_4 = Enforcement_handled_data['Enforcement_4']

    title = Enforcement_1_uncut.columns[0]

    Enforcement_headline_dict = {
        'Enforcement_cutoff': re.search('\d{1,2} \w+ \d{4}', title).group(),
        'Enforcement_total': Enforcement_1.loc[0, 'Current Month'],
        'Enforcement_total_line': Change_line_in_DR(Enforcement_1.loc[0, 'Change'])
    }
    
    Enforcement_section_dict = {
        'Enforcement_cutoff': re.search('\d{1,2} \w+ \d{4}', title).group(),
        'Enforcement_total': Enforcement_1.loc[0, 'Current Month'],
        'Enforcement_total_line': Change_line_in_DR(Enforcement_1.loc[0, 'Change']),
        'Enforcement_JIT_building_total': Enforcement_4.iloc[0, 4],
        'Enforcement_JIT_inspection_total': Enforcement_4.loc[0, 'Current Month'],
        'Enforcement_JIT_inspection_total_line': Change_line_in_DR(Enforcement_4.loc[0, 'Change']),
        'Enforcement_HHSRS_cat_1': Enforcement_2.iloc[0, 1],
        'Enforcement_HHSRS_cat_2': Enforcement_2.iloc[0, 2],
        'Enforcement_improvement_notices': Enforcement_3.iloc[0, 1],
        'Enforcement_hazard_awareness_notices': Enforcement_3.iloc[0, 2],
        'Enforcement_prohibition_order': Enforcement_3.iloc[0, 3],
        'Enforcement_improvement_notice_appeals': Enforcement_3.iloc[0, 5]
    }

    return Enforcement_headline_dict, Enforcement_section_dict