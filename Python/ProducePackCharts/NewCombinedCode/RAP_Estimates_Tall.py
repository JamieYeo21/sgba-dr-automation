"""
Created on Wednesday 28 August 2025, 09:23:10

author: Matthew Bandura
"""

import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np
from matplotlib.patches import FancyArrowPatch

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Navigate to the Utility folder relative to the script
utility_path = os.path.join(script_dir, 'Utility')

# Add it to sys.path so that python can import from it
sys.path.append(utility_path)

# Now you can import your functions
from Utility.functions import chop_df, get_excel_path

def create_18m_RAP_estimates(type, figure_count, colours, grey, secondary_grey, paths_variables, data_label_font_dict_white, data_label_font_dict_black):
    ###########
    # Main script Notifications
    if type==0:
        print(f'Figure{figure_count}_18m_RAP_estimates')

    if type==1:
        print(f'Accessible_Figure{figure_count}_18m_RAP_estimates')
    ###########


    ###########
    # CREATING THE DF
    ###########
    # Accessing the folder which stores the MI tables

    MI_tables_path = paths_variables['MI_tables_path']
    partial_output_path= paths_variables['partial_output_path']


    Combined_8 = pd.read_excel(MI_tables_path, sheet_name='Combined_8')
    Combined_8 = chop_df(Combined_8, 3, 4)

    Estimated_4 = pd.read_excel(MI_tables_path, sheet_name='Estimated_4')
    Estimated_4 = chop_df(Estimated_4, 3, 4)

    # Select the required data
    no_18m_buildings = Combined_8.iloc[:, 3].reset_index(drop=True)
    est_no_18m_buildings = Estimated_4.loc[[1]]

    max_total = no_18m_buildings[3] 

    yet_to_identify_lower = est_no_18m_buildings.iloc[0,1] - max_total
    yet_to_identify_upper =  est_no_18m_buildings.iloc[0,2] - max_total - yet_to_identify_lower
   

    data = pd.DataFrame({
        "Complete": [no_18m_buildings[0]],
        "Identified - remediation underway": [no_18m_buildings[1]],
        "Identified - in programme": [no_18m_buildings[2]],
        "Estimated yet to identify - low estimate": [yet_to_identify_lower],
        "Estimated yet to identify - high estimate": [yet_to_identify_upper]
    }, index=["Number of buildings"])


    ###########
    # CREATING THE GRAPH
    ###########
    fig, ax = plt.subplots(figsize=(18, 3), constrained_layout=True)
    bottom = np.zeros(len(data.index))
    colours = colours[:3]
    colours.append(secondary_grey)
    colours.append(grey)

    for i in range(data.shape[1]):
        bars = ax.barh(data.index, data.iloc[:, i], left=bottom, color=colours[i], label=data.columns[i], height = 0.5)
        bottom += data.iloc[:, i].astype(float)

        for j, bar in enumerate(bars):
            width = bar.get_width()
            if width <= 0:
                continue

            bar_color = mcolors.to_rgb(colours[i])
            luminance = 0.2126 * bar_color[0] + 0.7152 * bar_color[1] + 0.0722 * bar_color[2]
            font_dict = data_label_font_dict_black if luminance > 0.5 else data_label_font_dict_white
            label = f'{width:.0f}'
         
            if i < 3:
                if width / max_total >= 0.019:
                    label = f'{width:.0f}'
                    ax.text(bar.get_x() + width / 2, bar.get_y() + bar.get_height() / 2, label, **font_dict)    
            elif i == 4:
                start = bar.get_x() - data.iloc[:, 3].values[0]
                end = bar.get_x() + bar.get_width()
                midpoint = (start + end) / 2

                combined_label = f"{data.iloc[:, 3].values[0]:.0f} – {data.iloc[:, 4].values[0] + data.iloc[:, 3].values[0]:.0f}"
                ax.text(midpoint, bar.get_y() + bar.get_height() / 2, combined_label, **font_dict)



    # Formatting
    ax.xaxis.tick_top()
    ax.spines['bottom'].set_color('white')
    ax.spines['top'].set_color('darkgrey') 
    ax.spines['right'].set_color('white')
    ax.spines['left'].set_color('darkgrey')
    
    # Add horizontal arrow
    arrow = FancyArrowPatch((max_total, -0.1), (max_total + yet_to_identify_upper + yet_to_identify_lower, -0.1),
                        arrowstyle='<->', color='black',
                        linewidth=2, mutation_scale=20)
    ax.add_patch(arrow)

    ax.tick_params(axis='x', colors='darkgrey')
    ax.set_axisbelow(True)
    ax.legend(fontsize = 12,
            loc='upper center', 
            bbox_to_anchor=(0.5, 1.45),
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
    plt.savefig(output_path)
    plt.close(fig)

    print('DONE!')
    figure_count += 1
    return figure_count