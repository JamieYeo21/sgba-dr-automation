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

def create_SocialHousing2_Height(type, figure_count, colours, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black):
    ###########
    # Main script Notifications
    if type==0:
        print(f'Figure{figure_count}_SocialHousing2_Height')

    if type==1:
        print(f'Accessible_Figure{figure_count}_SocialHousing2_Height')
    ###########


    ###########
    # CREATING THE DF
    ###########
    # Accessing the folder which stores the MI tables
    MI_folder_path = MI_folder_path
    MI_tables_path = get_excel_path(MI_folder_path)

    # Accessing and transforming Combined_2
    Social_1 = pd.read_excel(MI_tables_path, sheet_name='Social_1')
    Social_1a = chop_df(Social_1, 3, 5)

    # Select the required columns
    number_of_11_18m = Social_1a.iloc[:, 1].reset_index(drop=True)
    number_of_18m = Social_1a.iloc[:, 3].reset_index(drop=True)

    max_total = max(sum(number_of_11_18m), sum(number_of_18m))

    data = pd.DataFrame({
        "Remediation complete": [number_of_11_18m[0],number_of_18m[0]],
        "Remediation complete - awaiting building control sign-off": [number_of_11_18m[1], number_of_18m[1]],
        "Remediation works in progress": [number_of_11_18m[2], number_of_18m[2]],
        "Remediation works planned": [number_of_11_18m[3], number_of_18m[3]],
        "Remediation plans unclear": [number_of_11_18m[4], number_of_18m[4]],
    }, index=["11-18m", "18m+"])


    ###########
    # CREATING THE GRAPH
    ###########
    fig, ax = plt.subplots(figsize=(12, 5))
    bottom = np.zeros(len(data.index))

    for i in range(data.shape[1]):
        bars = ax.barh(data.index, data.iloc[:, i], left=bottom, color=colours[i], label=data.columns[i])
        bottom += data.iloc[:, i].astype(float)

        for j, bar in enumerate(bars):
            width = bar.get_width()
            if width <= 0:
                continue

            bar_color = mcolors.to_rgb(colours[i])
            luminance = 0.2126 * bar_color[0] + 0.7152 * bar_color[1] + 0.0722 * bar_color[2]
            font_dict = data_label_font_dict_black if luminance > 0.5 else data_label_font_dict_white
            label = f'{width:.0f}'
         
            if width / max_total >= 0.019:
                label = f'{width:.0f}'
                ax.text(bar.get_x() + width / 2, bar.get_y() + bar.get_height() / 2, label, **font_dict)

    # Add total labels at the end of each bar
    totals = data.sum(axis=1).values
    for i, total in enumerate(totals):
        ax.text(
            total + 0.005*max_total,  
            i,          
            f'{total:.0f}',
            fontname='Arial',
            fontsize=12,
            fontweight='bold',
            va='center',
            ha='left',
            color='black'
    )

    # Formatting
    ax.xaxis.tick_top()
    ax.spines['bottom'].set_color('white')
    ax.spines['top'].set_color('darkgrey') 
    ax.spines['right'].set_color('white')
    ax.spines['left'].set_color('darkgrey')
    ax.tick_params(axis='x', colors='darkgrey')
    ax.set_axisbelow(True)
    ax.legend(fontsize = 12,
            loc='upper center', 
            bbox_to_anchor=(0.5, 1.5),
            fancybox=False,
            shadow=False,
            ncol=2,
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