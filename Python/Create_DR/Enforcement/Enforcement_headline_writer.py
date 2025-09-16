# Enforcement_writer_headline.py
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

def Enforcement_headline_writer(Enforcement_headline_dict, dates_variables, DR):
    # Unpack date vaiables
    last_month = dates_variables['last_month']

    # Unpack variables
    Enforcement_cutoff = Enforcement_headline_dict['Enforcement_cutoff']
    Enforcement_total = Enforcement_headline_dict['Enforcement_total']
    Enforcement_total_line = Enforcement_headline_dict['Enforcement_total_line']

    text = f'Enforcement – monthly update (as at {Enforcement_cutoff}) since previous publication'
    paragraph = DR.add_paragraph(text, style = 'Heading 3')

    # Paragraph
    text = f'As at {Enforcement_cutoff}, local authority enforcement action has been, or is being, taken under the Housing Act 2004 against {Enforcement_total} buildings over 11m with suspected unsafe cladding, {Enforcement_total_line} since the end of {last_month}.'
    DR.add_paragraph(text, style = 'Normal')