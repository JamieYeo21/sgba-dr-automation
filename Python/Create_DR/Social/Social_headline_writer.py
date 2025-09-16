# Social_writer_headline.py
"""
Created on Thursday 29 May 2025, 14:38:00

Author: Harry Simmons
"""

from docx import Document
from docx.enum.text import WD_COLOR_INDEX


def Social_headline_writer(DR):
    paragraph = DR.add_paragraph(style = 'Heading 1')
    run = paragraph.add_run('[ADD SOCIAL HEADLINE HERE]')
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    run.bold = True