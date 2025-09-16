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

# Add the adjacent folder to sys.path
folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))
sys.path.append(folder_path)

# Now you can import your functions
from Utility.functions import chop_df, get_excel_path
from Utility.MakeCurlyBrace import curlyBrace

def create_SocialHousing_Remediation3_Curly(type, figure_count, colours, grey, parital_output_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict):
    ###########
    # Main script Notifications
    if type==0:
        print(f'Figure{figure_count}_SocialHousing_Remediation3_Curly')

    if type==1:
        print(f'Accessible_Figure{figure_count}_SocialHousing_Remediation3_Curly')
    ###########


    ###########
    # CREATING THE DF
    ###########
    # Accessing the folder which stores the MI tables
    folder_path = 'Q:\\BSP\\Automation\\DR Automation\\Excel_inputs\\[PUT MI TABLES HERE]'
    MI_tables_path = get_excel_path(folder_path)

    # Accessing and transforming Combined_2
    Social_1 = pd.read_excel(MI_tables_path, sheet_name='Social_1')
    Social_1a = chop_df(Social_1, 3, 5)

    # Select the required columns
    number_of_buildings = Social_1a.iloc[:, 5].reset_index(drop=True)

    total = sum(number_of_buildings)

    yet_to_be_completed = number_of_buildings[1:].sum()
 
    data = pd.DataFrame({
        "No plans in place": [0] + [number_of_buildings[4]] * (len(number_of_buildings) - 1),
        "Plans in place": [0] + [number_of_buildings[3]] * (len(number_of_buildings) - 1),
        "Remediation started": [0] + [number_of_buildings[2]] * (len(number_of_buildings) - 1),
        "Remediation complete - awaiting building control sign-off": [yet_to_be_completed] + [number_of_buildings[1]] * (len(number_of_buildings) - 1),
        "Remediation complete": [number_of_buildings[0]] * len(number_of_buildings)
    }, index=["Remediation complete", "Remediation complete - awaiting\nbuilding control sign-off", "Remediation started", "Plans in place", "No plans in place"])


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

            # Skip white bars
            if color[:3] == mcolors.to_rgb("white"):
                continue

            # Determine font color based on luminance
            bar_color = color[:3]
            luminance = 0.2126 * bar_color[0] + 0.7152 * bar_color[1] + 0.0722 * bar_color[2]
            font_dict = data_label_font_dict_black if luminance > 0.5 else data_label_font_dict_white

            # Labels offset
            data_label_offset = 0.005

            stack_base = bottom[j]
            stack_top = stack_base + height

            if height == 0:
                continue  # skip zero-height bars

            if height < 0.032 * total:
                # Small bar: label above the stack
                data_y = stack_top + data_label_offset * total
                ax.text(bar.get_x() + bar.get_width() / 2, data_y,
                        f'{int(height)}', ha='center', va='bottom', **data_label_font_dict_black)
            else:
                # Normal bar: label inside
                data_y = bar.get_y() + height / 2
                ax.text(bar.get_x() + bar.get_width() / 2, data_y,
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
    ax.set_xticks(range(len(data.index)))
    ax.set_xticklabels(["Total buildings"] + data.index.tolist()[1:], fontsize=12)

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

    ax.text(x=-0.5,
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
        output_path = f'{parital_output_path}Figure{figure_count}.svg'

    if type==1:
        output_path = f'{parital_output_path}Accessible_Figure{figure_count}.svg' 

    plt.xticks(rotation=0)
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close(fig)

    print('DONE!')
    figure_count += 1
    return figure_count