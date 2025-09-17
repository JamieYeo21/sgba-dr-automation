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
from Utility.MakeCurlyBrace import curlyBrace

def create_CSS_Eligibility2(type, figure_count, colours, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict):
    ###########
    # Main script Notifications
    if type==0:
        print(f'Figure{figure_count}_CSS_Eligibility2')

    if type==1:
        print(f'Accessible_Figure{figure_count}_CSS_Eligibility2')
    ###########


    ###########
    # CREATING THE DF
    ###########
    # Accessing the folder which stores the MI tables
    MI_folder_path = MI_folder_path
    MI_tables_path = get_excel_path(MI_folder_path)

    # Accessing and transforming Combined_2
    CSS_1 = pd.read_excel(MI_tables_path, sheet_name='CSS_1')
    CSS_1a = chop_df(CSS_1, 3, 2)
    CSS_1b = chop_df(CSS_1, 8, 3)

    # Select the required columns
    number_of_pre_eligible = CSS_1a.iloc[:, 1].reset_index(drop=True)
    number_of_eligible = CSS_1b.iloc[:, 1].reset_index(drop=True)

    max_total = max(sum(number_of_pre_eligible), sum(number_of_eligible))
    max_pre_eligible = max(number_of_pre_eligible)
    pre_eligible_total = sum(number_of_pre_eligible)
    max_eligible = max(number_of_eligible)
    eligible_total = sum(number_of_eligible)

    # calculates where to place the brace
    pre_eligible_brace_x_value = max_pre_eligible + (0.005 * max_total)
    eligible_brace_x_valuee = max_eligible + (0.005 * max_total)
    pre_eligible_label_x_value = max_pre_eligible + (0.02 * max_total)
    eligible_label_x_value = max_eligible + (0.02 * max_total)

    data = pd.DataFrame({
        "Remediation complete": [number_of_eligible[2], number_of_eligible[1], number_of_eligible[0], number_of_pre_eligible[1], number_of_pre_eligible[0]]
    }, index=["Works completed", "Works started", "Works not started", "Live application", "Pre - application"])

    
    ###########
    # CREATING THE GRAPH
    ###########
    fig, ax = plt.subplots(figsize=(12, 5))
    bottom = np.zeros(len(data.index))

    for i in range(data.shape[1]):
        bars = ax.barh(data.index, data.iloc[:, i], left=bottom, color=colours[i], label=data.columns[i])
        
        for j, bar in enumerate(bars):
            width = bar.get_width()
            if width <= 0:
                continue

            bar_color = mcolors.to_rgb(colours[i])
            luminance = 0.2126 * bar_color[0] + 0.7152 * bar_color[1] + 0.0722 * bar_color[2]
            font_dict = data_label_font_dict_black if luminance > 0.5 else data_label_font_dict_white

            if width / max_total >= 0.019:
                label = f'{width:.0f}'
                ax.text(
                    bar.get_x() + width / 2,
                    bar.get_y() + bar.get_height() / 2,
                    label,
                    ha='center',
                    va='center',
                    **font_dict
                )
                data.iloc[j, i] = -1

        bottom += data.iloc[:, i].astype(float)

    # Add total labels at the end of each bar
    totals = data.sum(axis=1).values
    for i, total in enumerate(totals):
        if all(data.iloc[i, :] != -1):
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
    ax.grid(which = 'major',
          axis = 'x',
          color = 'darkgrey',
          linewidth = 0.3
          )

    # adding brace for Pre-eliglbe
    curlyBrace(fig,
               ax, 
               (pre_eligible_brace_x_value,4.4),
               (pre_eligible_brace_x_value,2.6),
               k_r = 0.02, 
               color = 'black',
               linewidth = 1,
               )

    # adding label for Pre-eliglbe
    ax.text(x=pre_eligible_label_x_value, 
            y=3.5, 
            s="Pre-eligible = " + str(pre_eligible_total), 
            ha='left',
            va='center',
            fontdict = brace_label_font_dict
            )

    # adding brace for Eliglbe buildings
    curlyBrace(fig,
               ax, 
               (eligible_brace_x_valuee,2.4)
               ,(eligible_brace_x_valuee,-0.4),
               k_r = 0.02, 
               color = 'black',
               linewidth = 1,
               )

    # adding label for Eliglbe buildings
    ax.text(x=eligible_label_x_value,
            y=1, 
            s="Eligible Buildings = " + str(eligible_total), 
            ha='left',
            va='center',
            fontdict = brace_label_font_dict
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