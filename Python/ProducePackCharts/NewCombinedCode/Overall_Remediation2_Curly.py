"""
Created on Thursday 20 February 2025, 11:27:10

author: Harry Simmons
"""

import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
from matplotlib.patches import Patch
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

def create_Overall_Remediation2_Curly(type, figure_count, colours, grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict):
    ###########
    # Main script Notifications
    if type==0:
        text = f'Figure{figure_count}_Overall_Remediation2_Curly'
        print(text)

    if type==1:
        text = f'Accessible_Figure{figure_count}_Overall_Remediation2_Curly'
        print(text)
    ###########


    ###########
    # CREATING THE DF
    ###########
    # Accessing the folder which stores the MI tables
    MI_folder_path = MI_folder_path
    MI_tables_path = get_excel_path(MI_folder_path)

    # Accessing and transforming Combined_2
    Combined_2 = pd.read_excel(MI_tables_path, sheet_name='Combined_2')
    Combined_2 = chop_df(Combined_2, 3, 3)

    # Select the required columns
    number_of_buildings = Combined_2.iloc[:, 5].reset_index(drop=True)

    total = sum(number_of_buildings)

    yet_to_be_completed = number_of_buildings[1:].sum()
 
    data = pd.DataFrame({
        "In programme": [0] + [number_of_buildings[2]] * (len(number_of_buildings) - 1),
        "Remediation underway": [yet_to_be_completed] + [number_of_buildings[1]] * (len(number_of_buildings) - 1),
        "Remediation complete": [number_of_buildings[0]] * len(number_of_buildings)
    }, index=["Remediation complete", "Remediation underway", "In programme"])


    ###########
    # CREATING THE GRAPH
    ###########
    fig, ax = plt.subplots(figsize=(13, 6))

    chopped_colours = colours[:len(number_of_buildings)]
    colours = chopped_colours[::-1]

    # Plotting the stacked bar chart
    bottom = np.zeros(len(data))
    for i, column in enumerate(data.columns):
        bar_colors = []
        for j in range(len(data)):
            if j == 0:
                # First bar: show all stacks
                bar_colors.append(grey if i == len(data.columns) - 2 else colours[i])
            elif i == len(data.columns) - j - 1:
                # Highlight only the corresponding stack for each bar
                bar_colors.append(colours[i])
            else:
                # All other segments white with no edge
                bar_colors.append("white")

        # Create bars and store them
        bars = ax.bar(data.index, data[column], bottom=bottom, color=bar_colors, width=0.5, edgecolor="none")
        
        # Add data labels
        for j, bar in enumerate(bars):
            color = bar.get_facecolor()
            height = bar.get_height()
            if color[:3] == mcolors.to_rgb("white") or height < 0.032 * total:
                continue  # Skip white bars

            bar_color = color[:3]  # RGB tuple
            luminance = 0.2126 * bar_color[0] + 0.7152 * bar_color[1] + 0.0722 * bar_color[2]
            font_dict = data_label_font_dict_black if luminance > 0.5 else data_label_font_dict_white

            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_y() + height / 2,
                    f'{int(height)}', ha='center', va='center', **font_dict)

        bottom += data[column]

    # Formatting
    ax.spines['bottom'].set_color('darkgrey')
    ax.spines['top'].set_color('None') 
    ax.spines['right'].set_color('None')
    ax.spines['left'].set_color('None')
    ax.tick_params(axis='x', colors='black', labelsize = 12)
    ax.tick_params(axis='y', colors='None')
    ax.yaxis.label.set_color('black')

    legend_names = ['Remediation complete'] + ['Yet to be completed']
    legend_colours = [colours[-1], grey]

    handles = [Patch(facecolor=color, edgecolor='none') for color in legend_colours]

    ax.legend(
        handles,
        legend_names,
        loc='upper right',
        edgecolor = 'None',
        facecolor = 'None',
        fontsize = 13,
        )
    
    curlyBrace(fig,
               ax, 
               (-0.25,0),
               (-0.25,total),
               k_r = 0.02, 
               color = 'black',
               linewidth = 1,
               )

    ax.text(x=-0.4,
            y=0.5 * total, 
            s=str(total), 
            ha='left',
            va='center',
            rotation = 'horizontal',
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