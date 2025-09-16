# DR_setup.py
"""
Created on Tuesday 17 December 2024, 14:59:41

Author: Harry Simmons
"""

import docx
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

def setup_doc():
    try:
        # Creating doc
        DR = docx.Document()

        # Setting margins
        sections = DR.sections
        for section in sections:
            section.left_margin = Cm(2)
            section.right_margin = Cm(2)

        # Setting a Heading 1 style
        style = DR.styles['Heading 1']
        font = style.font
        font.name = 'Aptos'
        font.size = Pt(24)
        font.bold = False
        font.color.rgb = RGBColor(0, 0, 0)
        paragraph_format = style.paragraph_format
        paragraph_format.space_after = Pt(12)

        # Setting a Heading 2 style
        style = DR.styles['Heading 2']
        font = style.font
        font.name = 'Aptos'
        font.size = Pt(18)
        font.bold = False
        font.color.rgb = RGBColor(0, 0, 0)
        paragraph_format = style.paragraph_format
        paragraph_format.space_after = Pt(12)

        # Setting a Heading 3 style
        style = DR.styles['Heading 3']
        font = style.font
        font.name = 'Aptos'
        font.size = Pt(13)
        font.bold = True
        font.color.rgb = RGBColor(0, 0, 0)
        paragraph_format = style.paragraph_format
        paragraph_format.space_after = Pt(12)

        # Setting a Normal style
        style = DR.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(12)
        font.bold = False
        font.color.rgb = RGBColor(0, 0, 0)
        paragraph_format = style.paragraph_format
        paragraph_format.space_after = Pt(12)

        return DR
    except Exception as e:
        print(f"An error occurred: {e}")
        return None