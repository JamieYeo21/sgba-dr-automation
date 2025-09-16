# dates.py
"""
Created on Wednesday 18 December 2024, 08:42:01

Author: Harry Simmons
"""

import re
import pandas as pd
import calendar as cal
from datetime import datetime
from Utility.functions import get_excel_path

# Function to create end of next quarter in words
def calculate_end_of_quarter_word(month, year):
    if month in [1, 2]:
        quarter_end_word = f'March {year}'
    elif month in [3, 4, 5]:
        quarter_end_word = f'June {year}'
    elif month in [6, 7, 8]:
        quarter_end_word = f'September {year}'
    elif month in [9, 10, 11]:
        quarter_end_word = f'December {year}'
    elif month == 12:
        quarter_end_word = f'March {year + 1}'
    else:
        return "Invalid month"
    
    return quarter_end_word

# Function to create end of next quarter in numbers
def calculate_end_of_quarter_no(month, year):
    if month in [1, 2]:
        next_quarter_start = datetime(year, 4, 1)
    elif month in [3, 4, 5]:
        next_quarter_start = datetime(year, 7, 1)
    elif month in [6, 7, 8]:
        next_quarter_start = datetime(year, 10, 1)
    elif month in [9, 10, 11]:
        next_quarter_start = datetime(year + 1, 1, 1)
    elif month == 12:
        next_quarter_start = datetime(year + 1, 4, 1)
    else:
        return "Invalid month"
    
    return next_quarter_start

def get_end_of_year_no(year, month):
    if month == 12:
        return datetime(year + 2, 1, 1)
    else:
        return datetime(year + 1, 1, 1)

def get_end_of_year_word(year, month):
    if month == 12:
        return f'December {year + 1}'
    else:
        return f'December {year}'

def extract_month_year(sheet_name, file_path):
    #Extracts month and year from the first column header of a given Excel sheet using regex searches.
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    title = df.columns[0]

    # Extract year
    year_match = re.search(r'\b\d{4}\b', title)
    year = int(year_match.group()) if year_match else None

    # Extract month
    month_match = re.search(r'\b(January|February|March|April|May|June|July|August|September|October|November|December)\b', title)
    month_word = month_match.group() if month_match else None

    # Convert month name to number
    month = {
        "January": 1, "February": 2, "March": 3, "April": 4,
        "May": 5, "June": 6, "July": 7, "August": 8,
        "September": 9, "October": 10, "November": 11, "December": 12
    }.get(month_word, None)
    
    #this ensures the dictionary returns appropriately named variables
    sheet_prefix = sheet_name[:3].lower()

    # Get last day of the month - monthrange returns the first weekday (ignored by _) and the final day (given to last_day)
    _, last_day = cal.monthrange(year, month)
    if sheet_name == 'Cover':
        return {
            "last_day": last_day,
            "year": year,
            "month_name": month_word,
            "month": month }
    else:
        return {
            f'{sheet_prefix}_last_day' : last_day,
            f'{sheet_prefix}_year' :  year,            
            f"{sheet_prefix}_month_name": month_word,
            f"{sheet_prefix}_month": month 
        }

def sort_dates():
    print('Handling Dates')
    # Accessing the folder which stores the MI tables
    folder_path = 'Q:\\BSP\\Automation\\DR Automation\\Excel_inputs\\[PUT MI TABLES HERE]'
    MI_tables_path = get_excel_path(folder_path)

    # Pull through the info about the next release date
    cover = pd.read_excel(MI_tables_path, sheet_name='Cover')
    publishing_cell_0 = cover.iloc[5,0]
    publishing_cell_1 = cover.iloc[6,0]

    #this is the info from the main DR
    cover_info = extract_month_year('Cover', MI_tables_path)
    month, month_name, year = cover_info['month'], cover_info['month_name'], cover_info['year']
    
    #this is the info related to the developer section of the DR
    dev_info = extract_month_year('Developer_3', MI_tables_path)
    dev_month, dev_year, dev_last_day = dev_info['dev_month'], dev_info['dev_year'], dev_info['dev_last_day']

    dev_cutoff = f"{dev_last_day} {cal.month_name[dev_month]} {dev_year}"

    # Working out dates for the main DR
    _, last_day = cal.monthrange(year, month)
    cutoff = f"{last_day} {cal.month_name[month]} {year}"
    last_month = cal.month_name[((month - 2) % 12 + 1)]
    this_month = f'{cal.month_name[month]} {year}'

    end_quarter_no = calculate_end_of_quarter_no(month, year)
    end_quarter_word = calculate_end_of_quarter_word(month, year)
    end_year_word = get_end_of_year_word(year, month)

    next_year = year + 1
    end_this_year = get_end_of_year_no(year, month)
    end_next_year = datetime(year + 2, 1, 1)
    last_year = year - 1

    last_year_month = f'{cal.month_name[month]} {last_year}'
    this_month_word = f'{cal.month_name[month]}'.lower()

    publishing_date_0 = (re.search(r'\d{1,2} \w+ \d{4}', publishing_cell_0)).group()
    publishing_date_1 = (re.search(r'\d{1,2} \w+ \d{4}', publishing_cell_1)).group()

    _this_month = f'{cal.month_name[month]}_{year}'

    if month == 1:
        last_month__year = year - 1
    else:
        last_month__year = year

    # Dictionary of variables for easy transport
    dates_variables = {
        'month': month,
        'year': year,
        'cutoff': cutoff,
        'last_month': last_month,
        'this_month': this_month,
        'hyperlink_month' : this_month.lower().replace(" ", "-"), #converts e.g. July 2025 to july-2025 for use in internal links in the DR (e.g. to the MI tables)
        'end_quarter_no': end_quarter_no,
        'end_quarter_word': end_quarter_word,
        'dev_month' : dev_month,
        'dev_year' : dev_year,
        'dev_last_day' : dev_last_day,
        'dev_cutoff' : dev_cutoff,
        'end_year_word': end_year_word,
        'next_year': next_year,
        'end_this_year': end_this_year,
        'end_next_year': end_next_year,
        'last_year_month': last_year_month,
        'this_month_word': this_month_word,
        'publishing_date_0':  publishing_date_0,
        'publishing_date_1':  publishing_date_1,
        'last_month_year': last_month__year,
        '_this_month':_this_month,
    }

    print('DONE!')
    return dates_variables

def hyperlink_convert(this_month):
    hyperlink_month = this_month.lower().replace(" ", "-")
    return hyperlink_month

