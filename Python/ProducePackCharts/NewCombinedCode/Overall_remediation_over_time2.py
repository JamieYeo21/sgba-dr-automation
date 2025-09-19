"""
Created on Thursday 20 February 2025, 11:27:10

author: Harry Simmons
"""

import re
import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Navigate to the Utility folder relative to the script
utility_path = os.path.join(script_dir, 'Utility')

# Add it to sys.path so that python can import from it
sys.path.append(utility_path)

# Now you can import your functions
from Utility.functions import chop_df, get_excel_path

def create_Overall_Remediation_over_time2(type, figure_count, colours, paths_variables, data_label_font_dict_white, data_label_font_dict_black):
    if type==0:
        text = f'Figure{figure_count}_Overall_remediation_over_time2'
        print(text)

    if type==1:
        text = f'Accessible_Figure{figure_count}_Overall_remediation_over_time2'
        print(text)

    # Accessing the folder which stores the MI tables
    MI_tables_path = paths_variables['MI_tables_path']
    partial_output_path = paths_variables['partial_output_path']

    # Pull through title
    cover = pd.read_excel(MI_tables_path, sheet_name='Cover')
    title = cover.columns[0]

    # Extract year using regex
    year_match = re.search(r'\b\d{4}\b', title)
    year = int(year_match.group()) if year_match else None

    # Extract month using regex
    month_match = re.search(r'\b(January|February|March|April|May|June|July|August|September|October|November|December)\b', title)
    month_word = month_match.group() if month_match else None

    # Convert month to its corresponding number
    month = {
        "January": 1, "February": 2, "March": 3, "April": 4,
        "May": 5, "June": 6, "July": 7, "August": 8,
        "September": 9, "October": 10, "November": 11, "December": 12
    }.get(month_word, None)

    # Accessing and transforming BSF_5
    Combined_6 = pd.read_excel(MI_tables_path, sheet_name='Combined_6')
    Combined_6 = chop_df(Combined_6, 4, 3)
    Combined_6 = Combined_6.iloc[:, -14:]

    # Generate headers
    headers = ["Remediation Stage"]
    for i in range(13):
        current_date = datetime(year, month, 1) - relativedelta(months=i)
        header = current_date.strftime('%b-%y')
        headers.append(header)
    headers[1:] = headers[1:][::-1]

    # Make headers strings
    headers = [str(header) for header in headers]

    # Assign headers to the DataFrame
    Combined_6.columns = headers

    remediation_stages = {
        "Remediation Stage": ['Complete', 'Underway', 'In Programme'],
    }

    # Replace the data in the first column
    for col in remediation_stages:
        Combined_6[col] = remediation_stages[col]

    # Convert the numeric columns to float
    Combined_6.iloc[:, 1:] = Combined_6.iloc[:, 1:].astype(float)  # <--- Added this line

    # Plotting the stacked bar chart
    fig, ax = plt.subplots(figsize = (12.5,6))
    bottom = np.zeros(len(Combined_6.columns) - 1, dtype=float)  # <--- Modified this line

    # Plot each stack in reverse order
    for i in range(Combined_6.shape[0] - 1, -1, -1):
        bars = ax.bar(Combined_6.columns[1:], Combined_6.iloc[i, 1:], bottom=bottom, color=colours[i], label=Combined_6.iloc[i, 0], width=0.6)
        bottom += Combined_6.iloc[i, 1:].astype(float)

        for j, bar in enumerate(bars):
            height = bar.get_height()
            bar_color = mcolors.to_rgb(colours[i]) # Convert color to RGB tuple
            luminance = 0.2126 * bar_color[0] + 0.7152 * bar_color[1] + 0.0722 * bar_color[2]

            # Choose font color based on luminance
            font_dict = data_label_font_dict_black if luminance > 0.5 else data_label_font_dict_white

            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_y() + height / 2,
                    f'{height:.0f}', **font_dict)

    # Add labels and title
    ax.spines['bottom'].set_color('darkgrey')
    ax.spines['top'].set_color('None') 
    ax.spines['right'].set_color('None')
    ax.spines['left'].set_color('None')
    ax.tick_params(axis='y', colors='None')
    ax.tick_params(axis='x', colors='black', labelsize = 15, rotation = 45)
    ax.yaxis.label.set_color('None')
    ax.xaxis.label.set_color('darkgrey')
    # Get handles and labels from the current legend
    handles, labels = ax.get_legend_handles_labels()  # <--- Added this line

    # Reverse the order of handles and labels
    handles = handles[::-1]  # <--- Added this line
    labels = labels[::-1]    # <--- Added this line

    # Create a new legend with the reversed order
    ax.legend(handles, labels, fontsize=13,  # <--- Modified this line
            loc='upper center',
            bbox_to_anchor=(1.1, 0.5),
            fancybox=False,
            shadow=False,
            ncol=1,
            edgecolor='None')
    fig.tight_layout()

    # Save the plot as SVG file
    if type==0:
        output_filename = f"Figure{figure_count}.svg"
        output_path = os.path.join(partial_output_path, output_filename)

    if type==1:
        output_filename = f"Accessible_Figure{figure_count}.svg"
        output_path = os.path.join(partial_output_path, output_filename)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close(fig)

    print('DONE!')
    figure_count += 1
    return figure_count