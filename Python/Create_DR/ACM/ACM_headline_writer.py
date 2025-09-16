# ACM_writer_headline.py
"""
Created on Monday 20 January 2025, 14:10:37

Author: Harry Simmons
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.enum.text import WD_COLOR_INDEX
import sys
import os

# Add the Utility folder to sys.path
folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))  # Replace 'folder_name' with the folder's name
sys.path.append(folder_path)

from ACM.ACM_variables import ACM_variable_creator

def ACM_headline_writer(ACM_headline_dict, dates_variables, DR):
    cutoff = dates_variables['cutoff']
    last_month = dates_variables['last_month']
    this_month = dates_variables['this_month']

    ACM_total = ACM_headline_dict['ACM_total']
    ACM_started_c_no = ACM_headline_dict['ACM_started_c_no']
    ACM_started_c_pct = ACM_headline_dict['ACM_started_c_pct']
    ACM_started_line = ACM_headline_dict['ACM_started_line']
    ACM_signoff_c_no = ACM_headline_dict['ACM_signoff_c_no']
    ACM_signoff_c_pct = ACM_headline_dict['ACM_signoff_c_pct']
    ACM_signoff_line = ACM_headline_dict['ACM_signoff_line']
    ACM_enforcement_total = ACM_headline_dict['ACM_enforcement_total']
    ACM_enforcement_pct = ACM_headline_dict['ACM_enforcement_pct']
    ACM_enforcement_line = ACM_headline_dict['ACM_enforcement_line']
    ACM_enforcement_vacants = ACM_headline_dict['ACM_enforcement_vacants']
    ACM_enforcement_forecast = ACM_headline_dict['ACM_enforcement_forecast']
    ACM_enforcement_enforced_no_forecast_word = ACM_headline_dict['ACM_enforcement_enforced_no_forecast_word']
    ACM_enforcement_not_enforced_no_forecast_word = ACM_headline_dict['ACM_enforcement_not_enforced_no_forecast_word']

    # Headline Title
    text = f'ACM remediation – monthly update (as at end {this_month}) since previous publication.'
    paragraph = DR.add_paragraph(text, style = 'Heading 3')

    # Paragraph 
    text = f'As at {cutoff} of the {ACM_total} high-rise (18 metres and over in height) residential and publicly owned buildings with ACM cladding systems, unlikely to meet Building Regulations, {ACM_started_c_no} ({ACM_started_c_pct}) have either started or completed remediation works, {ACM_started_line} since the end of {last_month}.'
    paragraph = DR.add_paragraph(text, style = 'Normal')

    # Paragraph
    text = f'Of these, {ACM_signoff_c_no} buildings ({ACM_signoff_c_pct}) have completed ACM remediation, including those awaiting building control sign-off, {ACM_signoff_line} since the end of {last_month}.'
    paragraph = DR.add_paragraph(text, style = 'Normal')

    # Paragraph
    paragraph = DR.add_paragraph(style = 'Normal')
    text = f'There are {ACM_enforcement_total} buildings yet to start ACM remediation ({ACM_enforcement_pct} of all buildings), {ACM_enforcement_line} since the end of {last_month}.'
    run = paragraph.add_run(text)
    text = ""
    if ACM_enforcement_vacants != 0:
        if ACM_enforcement_vacants == 'one':
            text += f' One building is vacant so does not pose a risk to resident safety'
        else:
            text += f' {ACM_enforcement_vacants.capitalize()} buildings are vacant so does not pose a risk to resident safety'
    if ACM_enforcement_forecast != 0:
        if ACM_enforcement_vacants == 'zero':
            if ACM_enforcement_forecast == 'one':
                text += f' One occupied building has a forecast start date'
            else:
                text += f' {ACM_enforcement_forecast.capitalize()} occupied buildings have forecast start dates'
        else:
            if ACM_enforcement_forecast == 'one':
                text += f', one occupied building has a forecast start date'
            else:
                text += f', {ACM_enforcement_forecast} occupied buildings have forecast start dates'
    if ACM_enforcement_enforced_no_forecast_word != 0:
        if ACM_enforcement_vacants == 'zero' and ACM_enforcement_forecast == 'zero':
            if ACM_enforcement_enforced_no_forecast_word == 'one':
                text += f' One building has had local authority enforcement action taken against them'
            else:
                text += f' {ACM_enforcement_enforced_no_forecast_word.capitalize()} buildings have had local authority enforcement action taken against them'
        else:
            if ACM_enforcement_enforced_no_forecast_word == 'one':
                text += f', a further building has had local authority enforcement action taken against them'
            else:
                text += f', {ACM_enforcement_enforced_no_forecast_word} further buildings have had local authority enforcement action taken against them'
    if ACM_enforcement_not_enforced_no_forecast_word != 0:
        if ACM_enforcement_not_enforced_no_forecast_word == 'one':
            text += f', the remaining building came into scope in 2024'
        else:
            text += f', the remaining {ACM_enforcement_not_enforced_no_forecast_word} buildings have come into scope in 2024 and 2025'

    text += '.'

    if ',' in text:
        parts = text.rsplit(',', 1)
        text = ' and'.join(parts)
    run = paragraph.add_run(text)