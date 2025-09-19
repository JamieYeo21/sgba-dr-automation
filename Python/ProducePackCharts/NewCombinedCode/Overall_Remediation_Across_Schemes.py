"""
Created on Thursday 20 February 2025, 11:27:10

author: Harry Simmons
"""

import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Navigate to the Utility folder relative to the script
utility_path = os.path.join(script_dir, 'Utility')

# Add it to sys.path so that python can import from it
sys.path.append(utility_path)

# Now you can import your functions
from Utility.functions import chop_df, get_excel_path

def create_Overall_Remediation_Across_Schemes(type, figure_count, colours, paths_variables, data_label_font_dict_white, data_label_font_dict_black):
    ###########
    # Main script Notifications
    if type==0:
        text = f'Figure{figure_count}_Overall_Remediation_Across_Schemes'
        print(text)

    if type==1:
        text = f'Accessible_Figure{figure_count}_Overall_Remediation_Across_Schemes'
        print(text)
    ###########


    ###########
    # CREATING THE DF
    ###########
    # Accessing the folder which stores the MI tables

    MI_tables_path = paths_variables['MI_tables_path']
    partial_output_path= paths_variables['partial_output_path']


    # Accessing, transforming and obtaining values from ACM_2
    ACM_2 = pd.read_excel(MI_tables_path, sheet_name='ACM_2')
    ACM_2 = chop_df(ACM_2, 2, 7)
    ACM_complete = ACM_2.iloc[0, 12] + ACM_2.iloc[1, 12]
    ACM_underway = ACM_2.iloc[2,12] + ACM_2.iloc[3,12]
    ACM_in_programme = ACM_2.iloc[4,12] + ACM_2.iloc[5,12] + ACM_2.iloc[6,12]

    # Accessing, transforming and obtaining values from BSF_1
    BSF_1 = pd.read_excel(MI_tables_path, sheet_name='BSF_1')
    BSF_1 = chop_df(BSF_1, 4, 5)
    BSF_complete = BSF_1.iloc[0, 10] + BSF_1.iloc[1, 10]
    BSF_underway = BSF_1.iloc[2,10]
    BSF_in_programme = BSF_1.iloc[3,10] + BSF_1.iloc[4,10]

    # Accessing, transforming and obtaining values from CSS_1
    CSS_1 = pd.read_excel(MI_tables_path, sheet_name='CSS_1')
    CSS_1 = chop_df(CSS_1, 8, 3)
    CSS_complete = CSS_1.iloc[2,2]
    CSS_underway = CSS_1.iloc[1,2]
    CSS_in_programme = CSS_1.iloc[0,2]

    # Accessing, transforming and obtaining values from Developer_1
    Developer_1 = pd.read_excel(MI_tables_path, sheet_name='Developer_1')
    Developer_1 = chop_df(Developer_1, 12, 5)
    Developer_complete = Developer_1.iloc[0, 6] + Developer_1.iloc[1, 6]
    Developer_underway = Developer_1.iloc[2, 6]
    Developer_in_programme = Developer_1.iloc[3, 6] + Developer_1.iloc[4, 6]

    # Accessing, transforming and obtaining values from Social_1
    Social_1 = pd.read_excel(MI_tables_path, sheet_name='Social_1')
    Social_1 = chop_df(Social_1, 3, 5)
    Social_complete = Social_1.iloc[0, 6] + Social_1.iloc[1, 6]
    Social_underway = Social_1.iloc[2, 6]
    Social_in_programme = Social_1.iloc[3, 6] + Social_1.iloc[4, 6]

    data = pd.DataFrame({
        "Remediation complete": [Social_complete, Developer_complete, CSS_complete, BSF_complete, ACM_complete],
        "Remediation underway": [Social_underway, Developer_underway, CSS_underway, BSF_underway, ACM_underway],
        "In programme": [Social_in_programme, Developer_in_programme, CSS_in_programme, BSF_in_programme, ACM_in_programme]
    }, index=["Social Housing", "Developer", "CSS", "BSF", "ACM"])


    ###########
    # CREATING THE GRAPH
    ###########


    # Normalize the data to 100%
    data_normalized = data.div(data.sum(axis=1), axis=0)

    # Create the horizontal 100% stacked bar chart
    fig, ax = plt.subplots(figsize=(10, 6))
    left = np.zeros(len(data_normalized))
    legend_labels = []

    for idx, column in enumerate(data_normalized.columns):
        label = column if column not in legend_labels else ""
        bars = ax.barh(data_normalized.index, data_normalized[column] * 100, left=left * 100, color=colours[idx], label=label)
        legend_labels.append(column)
        left += data_normalized[column]

    # Add data labels
    bottom = np.zeros(len(data_normalized))

    for i in range(data_normalized.shape[1]): 
        label = data_normalized.columns[i] if data_normalized.columns[i] not in legend_labels else None
        bars = ax.barh(data_normalized.index, data_normalized.iloc[:, i] * 100, left=bottom * 100, color=colours[i], label=label)
        legend_labels.append(data_normalized.columns[i])
        bottom += data_normalized.iloc[:, i].astype(float)

        for j, bar in enumerate(bars):
            width = bar.get_width()  
            if width < 3:
                    continue
            bar_color = mcolors.to_rgb(colours[i])
            luminance = 0.2126 * bar_color[0] + 0.7152 * bar_color[1] + 0.0722 * bar_color[2]
            font_dict = data_label_font_dict_black if luminance > 0.5 else data_label_font_dict_white
            label = f'{width:.0f}%'
            ax.text(bar.get_x() + width / 2, bar.get_y() + bar.get_height() / 2, label, **font_dict)

    # Formatting
    ax.xaxis.tick_top()
    ax.spines['bottom'].set_color('white')
    ax.spines['top'].set_color('darkgrey') 
    ax.spines['right'].set_color('white')
    ax.spines['left'].set_color('darkgrey')
    ax.tick_params(axis='x', colors='darkgrey')
    ax.set_axisbelow(True)
    ax.set_xlim(0, 100)
    ax.legend(fontsize = 12,
            loc='upper center', 
            bbox_to_anchor=(0.5, 1.15),
            fancybox=False,
            shadow=False,
            ncol=3,
            edgecolor = 'white'
            )  
    ax.grid(which = 'major',
          axis = 'x',
          color = 'darkgrey',
          linewidth = 0.3
          )


    ##########
    # SAVING THE GRAPH
    ##########

    # Save the plot as SVG file
    if type==0:
        output_filename = f"Figure{figure_count}.svg"
        output_path = os.path.join(partial_output_path, output_filename)

    if type==1:
        output_filename = f"Accessible_Figure{figure_count}.svg"
        output_path = os.path.join(partial_output_path, output_filename)
    plt.xticks(rotation=0)
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close(fig)

    print('DONE!')
    figure_count += 1
    return figure_count