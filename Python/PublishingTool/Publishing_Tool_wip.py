# Publishing_Tool.py
"""
Created on Thursday 21 August 2025, 11:52:00

Author: Matthew Bandura
"""

import os
import re
from docx import Document
from spire.doc import *
from spire.doc.common import *

print('RUNNING...')

def get_word_path(folder_path):
    files = os.listdir(folder_path)
    word_files = [file for file in files if file.endswith('.docx') or file.endswith('.doc')]
    if len(word_files) == 1:
        word_path = os.path.join(folder_path, word_files[0])
        return word_path
    else:
        raise FileNotFoundError("No Word document found in the folder.")
    

def extract_month_year_from_filename(filename):
    # Remove extension
    name_without_ext = os.path.splitext(filename)[0]

    # Define regex for month and year
    month_pattern = r'\b(January|February|March|April|May|June|July|August|September|October|November|December)\b'
    year_pattern = r'\b(20\d{2})\b'

    month_match = re.search(month_pattern, name_without_ext, re.IGNORECASE)
    year_match = re.search(year_pattern, name_without_ext)

    if month_match and year_match:
        month = month_match.group(1).capitalize()
        year = year_match.group(1)
        return month, year
    else:
        month = 'NO'
        year = 'DATE'
        return month, year


def word_to_md(input_folder, output_folder):
    word_path = get_word_path(input_folder)
    filename = os.path.basename(word_path)
    month, year = extract_month_year_from_filename(filename)

    # Create dynamic markdown path
    md_filename = f"publishing_markdown_{month}_{year}.md"
    md_path = os.path.join(output_folder, md_filename)
    # Create a Document object
    document = Document(word_path)
    fig_count = 1

    #remove the figures from the word doc before converting to markdown
    for i in range(document.Sections.Count):

        # Get a specific section
        section = document.Sections.get_Item(i)

        # Iterate through the paragraphs
        for j in range(section.Body.Paragraphs.Count):

            # Get a specific paragraph
            paragraph = section.Body.Paragraphs.get_Item(j)

            k = 0
            # Iterate through the child objects within the paragraph
            while k < len(paragraph.ChildObjects):
                
                # Get a specific paragraph
                obj = paragraph.ChildObjects.get_Item(k)

                # Determine if a child object is an image
                if isinstance(obj, DocPicture):
                        
                    # Remove the DocPicture instance
                    paragraph.ChildObjects.Remove(obj)

                    #replace the picture at k with text indicating the figure
                    paragraph.ChildObjects.Insert(k, paragraph.AppendText(f"![Figure {fig_count}](Figure{fig_count}.svg)"))
                    fig_count += 1
    

                #check if object found is an email    
                elif isinstance(obj, TextRange):
                    text = obj.Text
                    # Convert any email found (first argument) to Markdown format (by wrapping round <>, second argument)
                    new_text = re.sub(r'(\b[\w\.-]+@[\w\.-]+\.\w{2,}\b)', r'<\1>', text)
                    obj.Text = new_text
                    k += 1
                else:
                    k += 1
    #remove comments
    document.Comments.Clear()
    document.SaveToFile(md_path, FileFormat.Markdown)
    
    # Dispose resources
    document.Dispose()        


word_to_md('Q:\\BSP\Automation\\DR Automation\\Excel_inputs\\[PUT DR HERE]',
    'Q:\\BSP\Automation\\DR Automation\\DR_outputs\\DR_markdown')

print('DONE!')