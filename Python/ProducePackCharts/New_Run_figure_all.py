"""
Created on Tuesday 11 February 2025, 10:53:31

Author: Harry Simmons
"""

import time
import os 

# Start the timer
start_time = time.time()

# Overall
from NewCombinedCode.Overall_remediation_over_time2 import create_Overall_Remediation_over_time2
from NewCombinedCode.Overall_Remediation2_Curly import create_Overall_Remediation2_Curly
from NewCombinedCode.Overall_Remediation_Across_Schemes import create_Overall_Remediation_Across_Schemes
from NewCombinedCode.Overall_Height import create_Overall_Height
from NewCombinedCode.Overall_Tenure import create_Overall_Tenure
# Map

# ACM
from NewCombinedCode.ACM_Remediation import create_ACM_Remediation
# Projections 

# Dec19
from NewCombinedCode.ACM_By_Tenure import create_ACM_By_Tenure

# BSF
from NewCombinedCode.BSF_Remediation_Curly import create_BSF_Remediation_Curly
from NewCombinedCode.BSF_Remediation_over_time import create_BSF_Remediation_over_time
from NewCombinedCode.BSF_Tenure import create_BSF_Tenure

# CSS
from NewCombinedCode.CSS_Eligibility2 import create_CSS_Eligibility2
from NewCombinedCode.CSS_Height import create_CSS_Height
from NewCombinedCode.CSS_Tenure import create_CSS_Tenure

# Developer
from NewCombinedCode.Developer_Remediation_Curly import create_Developer_Remediation_Curly
from NewCombinedCode.Developer_Height import create_Developer_Height

# RAP
#from NewCombinedCode.RAP_Estimates_Tall import create_18m_RAP_estimates
#from NewCombinedCode.RAP_Estimates_Overall import create_11m_RAP_estimates

# Social
from NewCombinedCode.SocialHousing_Remediation3_Curly import create_SocialHousing_Remediation3_Curly
from NewCombinedCode.SocialHousing2_Height import create_SocialHousing2_Height


non_accessible_colour_scheme = ['#548235', '#A9D18E', '#E2F0D9', '#FFFF99', '#FFCC00', '#FF9900', '#FF0000']
non_accessible_grey = "#D4D4D4"
non_accessible_secondary_grey = "#979797"
accessible_colour_scheme = ['#2B4215', '#659B2C', '#333366', '#8687C1', '#FFFF00', '#E0612E', '#800000']
accessible_grey = "#D4D4D4"
accessible_secondary_grey = "#979797"
dept_colour = '#00625E'

non_accessible_figure_count = 1
accessible_figure_count = 1

# Get the current user's home directory (e.g., D:\Users\your.username)
home_dir = os.path.expanduser("~")

# Build the path to the graph output folder
partial_output_path = os.path.join(home_dir, "GitHub", "sgba-DR-automation", "DR_outputs", "DR_graphs")

# Build the path to the MI tables
MI_folder_path = os.path.join(
    home_dir,  "GitHub", "sgba-DR-automation", "Excel_inputs","[PUT MI TABLES HERE]"
)


# Ensure the directories exist
os.makedirs(partial_output_path, exist_ok=True)
os.makedirs(MI_folder_path, exist_ok=True)

data_label_font_dict_white = {
  "fontname": 'Arial',
  "fontsize": 12,
  "fontweight": 'bold',
  "horizontalalignment": 'center',
  "verticalalignment": 'center',
  "color": 'white'
  }

data_label_font_dict_black = {
  "fontname": 'Arial',
  "fontsize": 12,
  "fontweight": 'bold',
  "horizontalalignment": 'center',
  "verticalalignment": 'center',
  "color": 'black'
  }

brace_label_font_dict = {
  "fontname": 'Arial',
  "fontsize": 12,
  "fontweight":'normal',
  "color": 'black'
  }

def create_graphs(non_accessible_colour_scheme, non_accessible_grey, accessible_colour_scheme, accessible_grey, non_accessible_figure_count, accessible_figure_count, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict):
  ##########
  # NON ACCESSIBLE
  ##########



  # Overall
  non_accessible_figure_count = create_Overall_Remediation_over_time2(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  non_accessible_figure_count = create_Overall_Remediation2_Curly(0, non_accessible_figure_count, non_accessible_colour_scheme, non_accessible_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)
  non_accessible_figure_count = create_Overall_Remediation_Across_Schemes(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  non_accessible_figure_count = create_Overall_Height(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  non_accessible_figure_count = create_Overall_Tenure(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  
  #map
  non_accessible_figure_count += 1

  # ACM
  non_accessible_figure_count = create_ACM_Remediation(0, non_accessible_figure_count, non_accessible_colour_scheme, non_accessible_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)
  #Figures 8&9
  non_accessible_figure_count += 1
  non_accessible_figure_count += 1
  non_accessible_figure_count = create_ACM_By_Tenure(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)

  # BSF
  non_accessible_figure_count = create_BSF_Remediation_Curly(0, non_accessible_figure_count, non_accessible_colour_scheme, non_accessible_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)
  non_accessible_figure_count = create_BSF_Remediation_over_time(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  non_accessible_figure_count = create_BSF_Tenure(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)

  # CSS
  non_accessible_figure_count = create_CSS_Eligibility2(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)
  non_accessible_figure_count = create_CSS_Height(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  non_accessible_figure_count = create_CSS_Tenure(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)

  # Developer
  non_accessible_figure_count = create_Developer_Remediation_Curly(0, non_accessible_figure_count, non_accessible_colour_scheme, non_accessible_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)
  non_accessible_figure_count = create_Developer_Height(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)

  # RAP
 # non_accessible_figure_count = create_18m_RAP_estimates(0, non_accessible_figure_count, non_accessible_colour_scheme, non_accessible_grey, non_accessible_secondary_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
 # non_accessible_figure_count = create_11m_RAP_estimates(0, non_accessible_figure_count, non_accessible_colour_scheme, non_accessible_grey, non_accessible_secondary_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)

  # Social
  non_accessible_figure_count = create_SocialHousing_Remediation3_Curly(0, non_accessible_figure_count, non_accessible_colour_scheme, non_accessible_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)
  non_accessible_figure_count = create_SocialHousing2_Height(0, non_accessible_figure_count, non_accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)


  ##########
  # ACCESSIBLE
  ##########

  # Overall
  accessible_figure_count = create_Overall_Remediation_over_time2(1, accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  accessible_figure_count = create_Overall_Remediation2_Curly(1, accessible_figure_count, accessible_colour_scheme, accessible_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)
  accessible_figure_count = create_Overall_Remediation_Across_Schemes(1, accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  accessible_figure_count = create_Overall_Height(1, accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  accessible_figure_count = create_Overall_Tenure(1, accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)

  #Accessible map
  accessible_figure_count += 1

  # ACM
  accessible_figure_count =  create_ACM_Remediation(1, accessible_figure_count, accessible_colour_scheme, accessible_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)

  #Accessible figures 8&9
  accessible_figure_count += 1
  accessible_figure_count += 1
  accessible_figure_count = create_ACM_By_Tenure(1, accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)

  # BSF
  accessible_figure_count = create_BSF_Remediation_Curly(1,accessible_figure_count, accessible_colour_scheme, accessible_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)
  accessible_figure_count = create_BSF_Remediation_over_time(1, accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  accessible_figure_count = create_BSF_Tenure(1,accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)

  # CSS
  accessible_figure_count = create_CSS_Eligibility2(1, accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)
  accessible_figure_count = create_CSS_Height(1, accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  accessible_figure_count = create_CSS_Tenure(1, accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)

  # Developer
  accessible_figure_count = create_Developer_Remediation_Curly(1, accessible_figure_count, accessible_colour_scheme, accessible_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)
  accessible_figure_count = create_Developer_Height(1, accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  
  # RAP
  #accessible_figure_count = create_18m_RAP_estimates(1, accessible_figure_count, accessible_colour_scheme, accessible_grey, accessible_secondary_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)
  #accessible_figure_count = create_11m_RAP_estimates(1, accessible_figure_count, accessible_colour_scheme, accessible_grey, accessible_secondary_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)

  # Social
  accessible_figure_count = create_SocialHousing_Remediation3_Curly(1, accessible_figure_count, accessible_colour_scheme, accessible_grey, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)
  accessible_figure_count = create_SocialHousing2_Height(1, accessible_figure_count, accessible_colour_scheme, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black)


# Call the main function
if __name__ == "__main__":
  create_graphs(non_accessible_colour_scheme, non_accessible_grey, accessible_colour_scheme, accessible_grey, non_accessible_figure_count, accessible_figure_count, partial_output_path, MI_folder_path, data_label_font_dict_white, data_label_font_dict_black, brace_label_font_dict)


# End the timer and print the elapsed time
end_time = time.time()
elapsed_time = end_time - start_time
print(f"Program executed in {elapsed_time:.2f} seconds")
print('Done! Have a nice day :)')