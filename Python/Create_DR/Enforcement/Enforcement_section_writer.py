# Enforcement_writer_section.py
"""
Created on Tuesday 04 March 2025, 08:43:31

Author: Harry Simmons
"""

import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn

def Enforcement_section_writer(Enforcement_section_dict, dates_variables, DR):
    # Unpack dates variables
    last_month = dates_variables['last_month']

    # Unpack variables
    Enforcement_cutoff = Enforcement_section_dict['Enforcement_cutoff']
    Enforcement_total = Enforcement_section_dict['Enforcement_total']
    Enforcement_total_line = Enforcement_section_dict['Enforcement_total_line']
    Enforcement_JIT_building_total = Enforcement_section_dict['Enforcement_JIT_building_total']
    Enforcement_JIT_inspection_total = Enforcement_section_dict['Enforcement_JIT_inspection_total']
    Enforcement_JIT_inspection_total_line = Enforcement_section_dict['Enforcement_JIT_inspection_total_line']
    Enforcement_HHSRS_cat_1 = Enforcement_section_dict['Enforcement_HHSRS_cat_1']
    Enforcement_HHSRS_cat_2 = Enforcement_section_dict['Enforcement_HHSRS_cat_2']
    Enforcement_improvement_notices = Enforcement_section_dict['Enforcement_improvement_notices']
    Enforcement_hazard_awareness_notices = Enforcement_section_dict['Enforcement_hazard_awareness_notices']
    Enforcement_prohibition_order = Enforcement_section_dict['Enforcement_prohibition_order']
    Enforcement_improvement_notice_appeals = Enforcement_section_dict['Enforcement_improvement_notice_appeals']

    # Section Title 
    paragraph = DR.add_paragraph('Enforcement', style = 'Heading 2')

    # Paragraph
    text = f'Information in this section is correct as at {Enforcement_cutoff} and shows a monthly update from the previous publication.'
    DR.add_paragraph(text, style = 'Normal')

    # Paragraph
    DR.add_paragraph('This section includes local authority enforcement action on buildings 11 metres or over in height. Up until June 2024 only enforcement action taken on buildings over 18m in height was reported on.', style = 'Normal')

    # Heading
    paragraph = DR.add_paragraph('Local authority enforcement action: key statistics', style = 'Heading 3')

    # Paragraph
    text = f'As at {Enforcement_cutoff}, enforcement action has been, or is being, taken under the Housing Act against {Enforcement_total} buildings over 11m with suspected unsafe cladding - {Enforcement_total_line} since the end of {last_month}. Of these, {Enforcement_JIT_building_total} buildings have had {Enforcement_JIT_inspection_total} inspections with Joint Inspection Team support – {Enforcement_JIT_inspection_total_line} inspections since the end of {last_month}.'
    DR.add_paragraph(text, style = 'Normal')

    # Paragraph
    text = f'Of the {Enforcement_total} buildings where a local authority had undertaken an inspection, we are aware that:'
    DR.add_paragraph(text, style = 'Normal')

    # Bullet point
    text = f'{Enforcement_HHSRS_cat_1} had a Category 1 HHSRS rating'
    DR.add_paragraph(text, style = 'List Bullet')

    # Bullet point
    text = f'{Enforcement_HHSRS_cat_2} had a Category 2 HHSRS rating'
    DR.add_paragraph(text, style = 'List Bullet')

    # Paragraph
    DR.add_paragraph('As several buildings have now had two or more inspections, from August 2024, only the most recent inspection’s category rating is reported on, as this is used to gauge the risks of the buildings in each Local Authority’s portfolio. Additional inspections are still reported as enforcement action.', style = 'Normal')

    # Paragraph
    text = f'Of the {Enforcement_total} buildings, we are aware that at least {Enforcement_improvement_notices} improvement notices, {Enforcement_hazard_awareness_notices} hazard awareness notices and {Enforcement_prohibition_order} prohibition order have been served to date. Some buildings may have received multiple notices. We understand that {Enforcement_improvement_notice_appeals} of the improvement notices have been subject to appeals.'
    DR.add_paragraph(text, style = 'Normal')