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

from Utility.functions import chop_df, get_excel_path
from Utility.MakeCurlyBrace import curlyBrace

def create_ACM_Remediation(type, figure_count, colours, grey, paths_variables, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict):
    ###########
    # Main script Notifications
    if type==0:
        print(f'Figure{figure_count}_ACM_Remediation')

    if type==1:
        print(f'Accessible_Figure{figure_count}_ACM_Remediation')
    ###########


    ###########
    # CREATING THE DF
    ###########
    # Accessing the folder which stores the MI tables
    MI_tables_path = paths_variables['MI_tables_path']
    partial_output_path= paths_variables['partial_output_path']


    # Accessing and transforming ACM_3
    ACM_2 = pd.read_excel(MI_tables_path, sheet_name='ACM_2')
    ACM_2 = chop_df(ACM_2, 2, 7)

    # Accessing and transforming ACM_3
    ACM_3 = pd.read_excel(MI_tables_path, sheet_name='ACM_3')
    ACM_3a = chop_df(ACM_3, 3, 6)
    ACM_3a["Total"] = ACM_3a.iloc[:, 1:6].sum(axis=1)

    # Select the required columns
    number_of_buildings = ACM_2.iloc[:,11].reset_index(drop=True)
    number_of_vacants = ACM_3a.iloc[:,6].reset_index(drop=True)

    total = sum(number_of_buildings)
    cladding_removed = sum(number_of_buildings[:3])
    cladding_remaining = sum(number_of_buildings[3:])
    yet_to_be_completed = number_of_buildings[1:].sum()
 
    data = pd.DataFrame({
        "Remediation plan unclear / awaiting further advice": [0] + [number_of_buildings[6]] * (len(number_of_buildings) - 1),
        "Reported an intent to remediate and are developing plans": [0] + [number_of_buildings[5]] * (len(number_of_buildings) - 1),
        "Remediation plan in place": [0] + [number_of_buildings[4]] * (len(number_of_buildings) - 1),
        "Remediation started": [0] + [number_of_buildings[3]] * (len(number_of_buildings) - 1),
        "Remediation started - cladding removed": [0] + [number_of_buildings[2]] * (len(number_of_buildings) - 1),
        "Works complete awaiting building control sign-off": [yet_to_be_completed] + [number_of_buildings[1]] * (len(number_of_buildings) - 1),
        "Remediation complete": [number_of_buildings[0]] * len(number_of_buildings)
    }, index=["Remediation complete", "Works complete\nawaiting building\ncontrol sign-off", "Remediation\nstarted - cladding\nremoved", "Remediation\nstarted", "Remediation plans\nin place", "Reported an intent\nto remediate and\nare developing\nplans", "Remediation plan\nunclear / awaiting\nfurther advice"])

    vacants_data = pd.DataFrame({
        "Number of Vacants": number_of_vacants
    })

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
            elif j == len(data) - 1 and i == 0:
                # Last bar: highlight the first stack
                bar_colors.append(colours[i])
            elif i == len(data.columns) - j - 1:
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
            
            # Always label the last bar (j == len(data) - 1), even if it's white
            if color[:3] == mcolors.to_rgb("white"):
                continue

            # Determine font color based on luminance
            bar_color = color[:3]
            luminance = 0.2126 * bar_color[0] + 0.7152 * bar_color[1] + 0.0722 * bar_color[2]
            font_dict = data_label_font_dict_black if luminance > 0.5 else data_label_font_dict_white

            # Labels offset
            data_vacant_label_offset = 0.038
            data_no_vacant_label_offset = 0.005
            vacant_label_offset = 0.035

            stack_base = bottom[j]
            stack_top = stack_base + height

            if height == 0:
                continue  # skip zero-height bars

            # Vacant label logic
            vacant_index = j - 1
            vacant_label = None
            if j >= 1 and vacant_index < len(vacants_data):
                vacant_count = vacants_data.loc[vacant_index, "Number of Vacants"]
                if vacant_count != 0:
                    vacant_label = f'(incl. {int(vacant_count)} vacant)'

            if height < 0.032 * total:
                # Small stack: data label above the stack
                # If there's a vacant label, push the data label up to make room

                if vacant_label:
                    # Push data label higher to make room for vacant label below
                    data_y = stack_top + data_vacant_label_offset * total
                    vacant_y = stack_top + vacant_label_offset * total  
                    ax.text(bar.get_x() + bar.get_width() / 2, data_y,
                            f'{int(height)}', ha='center', va='bottom', **data_label_font_dict_black)
                    ax.text(bar.get_x() + bar.get_width() / 2, vacant_y,
                            vacant_label, ha='center', va='top', fontsize=10, color='black')
                else:
                    # No vacant label, place data label just above the stack
                    data_y = stack_top + data_no_vacant_label_offset * total
                    ax.text(bar.get_x() + bar.get_width() / 2, data_y,
                            f'{int(height)}', ha='center', va='bottom', **data_label_font_dict_black)

            else:
                # Normal stack: data label inside, vacant label above
                data_y = bar.get_y() + height / 2
                vacant_y = stack_top + vacant_label_offset * total  
                ax.text(bar.get_x() + bar.get_width() / 2, data_y,
                        f'{int(height)}', ha='center', va='center', **font_dict)
                if vacant_label:
                    ax.text(bar.get_x() + bar.get_width() / 2, vacant_y,
                            vacant_label, ha='center', va='bottom', fontsize=10, color='black')

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
    
    # Adding the brace the Total number of buildings
    curlyBrace(fig,
               ax, 
               (-0.25,0),
               (-0.25,total),
               k_r = 0.02, 
               color = 'black',
               linewidth = 1,
               )

    ax.text(x=-0.56,
            y=0.5 * total, 
            s=str(total), 
            ha='left',
            va='center',
            rotation = 'horizontal',
            fontdict = brace_label_font_dict
            )
    
    # Adding the brace for cladding removed
    curlyBrace(fig,
               ax, 
               (2.25,total),
               (2.25,cladding_remaining),
               k_r = 0.02, 
               color = 'black',
               linewidth = 1,
               )

    ax.text(x=2.4,
            y=cladding_remaining + 0.5 * cladding_removed, 
            s='Cladding removed: ' + str(cladding_removed), 
            ha='left',
            va='center',
            rotation = 'horizontal',
            fontdict = brace_label_font_dict
            )
    
    # Adding the brace for clading remaining
    curlyBrace(fig,
               ax, 
               (6.25,cladding_remaining),
               (6.25,0),
               k_r = 0.02, 
               color = 'black',
               linewidth = 1,
               )

    ax.text(x=6.3,
            y=0.5 * cladding_remaining, 
            s='Cladding\nremaining:\n' + str(cladding_remaining), 
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