"""
Created on Wednesday August 28 2025, 09:21:10

authors: Harry Simmons + Matthew Bandura
"""

import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np
from matplotlib.patches import FancyArrowPatch

# Add the adjacent folder to sys.path
folder_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'Utility'))
sys.path.append(folder_path)

# Now you can import your functions
from Utility.functions import chop_df, get_excel_path

def create_11m_RAP_estimates(type, figure_count, colours, grey, secondary_grey, parital_output_path, data_label_font_dict_white, data_label_font_dict_black):
    ###########
    # Main script Notifications
    if type==0:
        print(f'Figure{figure_count}_Overall_Height')

    if type==1:
        print(f'Accessible_Figure{figure_count}_Overall_Height')
    ###########


    ###########
    # CREATING THE DF
    ###########
    # Accessing the folder which stores the MI tables
    folder_path = 'Q:\\BSP\\Automation\\DR Automation\\Excel_inputs\\[PUT MI TABLES HERE]'
    MI_tables_path = get_excel_path(folder_path)

    # Accessing and transforming Combined_2 for the total nmber
    Combined_2 = pd.read_excel(MI_tables_path, sheet_name='Combined_2')
    Combined_2 = chop_df(Combined_2, 3, 4)
    
    # Accessing and transforming Estimated_2 for the estimates of total number off 11m+ buildings
    Estimated_2 = pd.read_excel(MI_tables_path, sheet_name='Estimated_2')
    Estimated_2 = chop_df(Estimated_2, 3, 4)

    # Select the required column
    no_11m_buildings = Combined_2.iloc[:, 5].reset_index(drop=True)
    est_no_11m_buildings = Estimated_2.loc[[2]]
    max_total = no_11m_buildings[3]
    yet_to_identify_lower = est_no_11m_buildings.iloc[0,1] - max_total
    yet_to_identify_upper = est_no_11m_buildings.iloc[0,2] - max_total - yet_to_identify_lower


    data = pd.DataFrame({
        "Complete": [no_11m_buildings[0]],
        "Identified - yet to complete": [no_11m_buildings[1] + no_11m_buildings[2]],
        "Estimated yet to identify - low estimate": [yet_to_identify_lower],
        "Estimated yet to identify - high estimate": [yet_to_identify_upper]
    }, index=["Number of buildings"])


    ###########
    # CREATING THE GRAPH
    ###########
    fig, ax = plt.subplots(figsize=(18, 3), constrained_layout=True)
    bottom = np.zeros(len(data.index))
    colours = colours[:2]
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
         
            if i < 2:
                if width / max_total >= 0.019:
                    label = f'{width:.0f}'
                    ax.text(bar.get_x() + width / 2, bar.get_y() + bar.get_height() / 2, label, **font_dict)    
            elif i == 3:
                start = bar.get_x() - data.iloc[:, 2].values[0]
                end = bar.get_x() + bar.get_width()
                midpoint = (start + end) / 2

                combined_label = f"{data.iloc[:, 2].values[0]:.0f} – {data.iloc[:, 3].values[0] + data.iloc[:, 2].values[0]:.0f}"
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
        output_path = f'{parital_output_path}Figure{figure_count}.svg'

    if type==1:
        output_path = f'{parital_output_path}Accessible_Figure{figure_count}.svg' 

    plt.xticks(rotation=0)
    plt.savefig(output_path)
    plt.close(fig)

    print('DONE!')
    figure_count += 1
    return figure_count