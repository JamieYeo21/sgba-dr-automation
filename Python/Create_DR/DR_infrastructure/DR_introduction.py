# DR_start_infrastructure.py
"""
Created on Thursday 13 March 2025, 10:05:24

Author: Harry Simmons
"""

import docx
from docx.shared import Pt, Cm, RGBColor
import sys
import os

folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))  # Replace 'folder_name' with the folder's name
sys.path.append(folder_path)

from Utility.functions import add_numbered_list_item

def DR_introduction(DR, figure_count):
    # Section Title 
    paragraph = DR.add_paragraph('Introduction', style = 'Heading 2')

    # Paragraph
    DR.add_paragraph('This Data Release provides data on:', style = 'Normal')
    
    #Bullet Point 1
    DR.add_paragraph('The estimated number of residential buildings 11 metres and over in height in England that have or had unsafe cladding requiring remediation.', style = 'List Bullet')

    #Bullet Point 2
    DR.add_paragraph('Remediation progress across MHCLG’s Building Safety Remediation portfolio, covering buildings in the ACM programme, Building Safety Fund, Cladding Safety Scheme, developer remediation contract and reported by registered providers of social housing.', style = 'List Bullet')

    #Bullet Point 3
    DR.add_paragraph('Progress with remediation of high-rise (18 metres and over in height) residential buildings (including student accommodation and hotels) and publicly owned buildings with ACM cladding systems unlikely to meet Building Regulations in the ACM programme.', style = 'List Bullet')

    #Bullet Point 4
    DR.add_paragraph('Progress with the applications for and remediation of medium-rise (11 – 18 metres in height) residential buildings in England, and Northern Ireland, and high-rise residential buildings outside of London with non-ACM cladding systems in the Cladding Safety Scheme.', style = 'List Bullet')

    #Bullet Point 5
    DR.add_paragraph('Progress of remediation of residential buildings 11 metres and over in height with life-critical fire safety risks under the developer remediation contract.', style = 'List Bullet')

    #Bullet Point 6
    DR.add_paragraph('Progress of remediation of residential buildings 11 metres and over in height with unsafe cladding reported by registered providers of social housing.', style = 'List Bullet')

    #Bullet Point 7
    DR.add_paragraph('Enforcement action taken by local authorities against high-rise residential buildings with unsafe cladding under the Housing Act 2004.', style = 'List Bullet')

    #Bullet Point 8
    DR.add_paragraph('In additional management information tables only, the progress of the Waking Watch Relief Fund and Waking Watch Replacement Fund. From June 2025 this includes data on the Waking Watch Replacement Fund 2023.', style = 'List Bullet')
    
    # Paragraph
    DR.add_paragraph(f'The data in Figure {figure_count} and the overall remediation progress section of the data release, shows the combined remediation progress across MHCLG’s Building Safety Remediation portfolio, covering buildings and accounting for crossover in the ACM programme, Building Safety Fund, Cladding Safety Scheme, developer remediation contract and reported by registered providers of social housing.', style = 'Normal')

    # Paragraph
    DR.add_paragraph('The figures in this publication are correct as at the specified dates. Remediation progress on the ACM programme, BSF and CSS will be updated monthly, and remediation progress on the developer remediation contract and registered providers of social housing will be updated quarterly.', style = 'Normal')

    # Paragraph
    DR.add_paragraph('From June 2025 the Building Safety Remediation data release includes estimates on the number of residential buildings 11 metres and over in height in England with unsafe cladding requiring remediation. These estimates are not expected to change frequently.', style = 'Normal')