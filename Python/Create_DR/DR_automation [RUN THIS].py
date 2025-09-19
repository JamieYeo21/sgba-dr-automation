# DR_automation.py
"""
Created on Wednesday 18 December 2024, 08:02:22

Author: Harry Simmons
"""

"""
This code takes all of the the individual sections of this projects and runs them in order here.
This was done to:
    - Make it easier to navigate this project by splitting up parts of it
"""

import time
import os 
# Start the timer
start_time = time.time()

from Utility.dates import sort_dates
from Utility.functions import create_paths

from DR_infrastructure.DR_setup import setup_doc
from DR_infrastructure.DR_start_infrastructure import DR_start
from DR_infrastructure.DR_introduction import DR_introduction
from DR_infrastructure.DR_enquiries import DR_enquiries
from DR_infrastructure.DR_building_safety_overview import DR_building_safety_overview
from DR_infrastructure.DR_end_infrastructure import DR_end

from Portfolio.Portfolio_data_handler import Portfolio_retrieve_data
from Portfolio.Portfolio_variables import Portfolio_variable_creator
from Portfolio.Portfolio_headline_writer import Portfolio_headline_writer
from Portfolio.Portfolio_section_writer import Portfolio_section_writer

from Estimates.Estimates_variables import Estimates_variable_creator
from Estimates.Estimates_section_writer import Estimates_section_writer

from ACM.ACM_data_handler import ACM_retrieve_data
from ACM.ACM_variables import ACM_variable_creator
from ACM.ACM_headline_writer import ACM_headline_writer
from ACM.ACM_section_writer import ACM_section_writer  

from BSF.BSF_data_handler import BSF_retrieve_data
from BSF.BSF_variables import BSF_variable_creator
from BSF.BSF_headline_writer import BSF_headline_writer
from BSF.BSF_section_writer import BSF_section_writer

from CSS.CSS_data_handler import CSS_retrieve_data
from CSS.CSS_variables import CSS_variable_creator
from CSS.CSS_headline_writer import CSS_headline_writer
from CSS.CSS_section_writer import CSS_section_writer

from Developer.Developer_data_handler import Developer_retrieve_data_last_month, Developer_retrieve_data_this_month
from Developer.Developer_variables import Developer_variable_creator
from Developer.Developer_headline_writer import Developer_headline_writer
from Developer.Developer_section_writer import Developer_section_writer

#from RAP.RAP_data_handler import RAP_retrieve_data
#from RAP.RAP_variables import RAP_variable_creator
#from RAP.RAP_section_writer import RAP_section_writer

from Social.Social_headline_writer import Social_headline_writer
from Social.Social_section_writer import Social_section_writer

from Enforcement.Enforcement_data_handler import Enforcement_retrieve_data
from Enforcement.Enforcement_variables import Enforcement_variable_creator
from Enforcement.Enforcement_headline_writer import Enforcement_headline_writer
from Enforcement.Enforcement_section_writer import Enforcement_section_writer


def create_DR():
    # Handle data
    figure_count = 1
    table_count = 1
    dates_variables = sort_dates()
    paths_variables = create_paths()

    Portfolio_handled_data = Portfolio_retrieve_data(paths_variables)
    ACM_handled_data = ACM_retrieve_data(dates_variables, paths_variables)
    BSF_handled_data = BSF_retrieve_data(paths_variables)
    CSS_handled_data = CSS_retrieve_data(paths_variables)
    Developer_handled_data_last_month = Developer_retrieve_data_last_month(paths_variables)
    Developer_handled_data_this_month = Developer_retrieve_data_this_month(paths_variables)
#    RAP_handled_data = RAP_retrieve_data(paths_variables)
    Enforcement_handled_data = Enforcement_retrieve_data(paths_variables)

    # Format variables
    print('Formatting Variables')
    Estimates_tables, Estimates_headline_dict, Estimates_section_dict = Estimates_variable_creator(Portfolio_handled_data)
    Portfolio_tables, Portfolio_headline_dict, Portfolio_section_dict = Portfolio_variable_creator(Portfolio_handled_data)
    ACM_tables, ACM_headline_dict, ACM_section_dict = ACM_variable_creator(ACM_handled_data, dates_variables)
    BSF_tables, BSF_headline_dict, BSF_section_dict, BSF_developer_transfers = BSF_variable_creator (BSF_handled_data)
    CSS_tables, CSS_headline_dict, CSS_section_dict = CSS_variable_creator (CSS_handled_data)
    Developer_tables, Developer_headline_dict, Developer_section_dict = Developer_variable_creator(Developer_handled_data_last_month, Developer_handled_data_this_month, dates_variables)
#    RAP_section_dict = RAP_variable_creator(RAP_handled_data)
    Enforcement_headline_dict, Enforcement_section_dict = Enforcement_variable_creator(Enforcement_handled_data)
    print('DONE!')

    # Write DR
    print('Writing Data Release')
    DR = setup_doc()
    DR_start(dates_variables, DR)
    figure_count = Portfolio_headline_writer(Portfolio_headline_dict, Estimates_headline_dict, figure_count, dates_variables, paths_variables, DR)
    ACM_headline_writer(ACM_headline_dict, dates_variables, DR)
    BSF_headline_writer(BSF_headline_dict, dates_variables, DR)
    CSS_headline_writer(CSS_headline_dict, dates_variables, DR)
    Developer_headline_writer(Developer_headline_dict, dates_variables, DR)
    Social_headline_writer(DR)
    Enforcement_headline_writer(Enforcement_headline_dict, dates_variables, DR)
    DR_introduction(DR, figure_count)
    DR_enquiries(dates_variables, DR)
    DR_building_safety_overview(DR)
    table_count = Estimates_section_writer(Estimates_section_dict, Estimates_tables, table_count, dates_variables, DR)
    figure_count, table_count= Portfolio_section_writer(Portfolio_section_dict, Portfolio_tables, figure_count, table_count, dates_variables, paths_variables, DR)
    figure_count, table_count = ACM_section_writer(ACM_section_dict, ACM_tables, figure_count, table_count, dates_variables, paths_variables, DR)
    figure_count, table_count = BSF_section_writer(BSF_section_dict, BSF_tables, figure_count, table_count, dates_variables, paths_variables, DR)
    figure_count, table_count = CSS_section_writer(CSS_section_dict, CSS_tables, figure_count, table_count, dates_variables, paths_variables, DR)
    figure_count, table_count = Developer_section_writer(Developer_section_dict, BSF_developer_transfers, Developer_tables, figure_count, table_count, dates_variables, paths_variables, DR)
#    figure_count = RAP_section_writer(RAP_section_dict,figure_count, dates_variables, DR)
    figure_count = Social_section_writer(DR, figure_count,paths_variables)
    Enforcement_section_writer(Enforcement_section_dict, dates_variables, DR)
    DR_end(DR, dates_variables)
    print('DONE!')

    # Saving the document
    this_month = dates_variables['this_month']
    save_path = paths_variables['save_path']

    output_path = os.path.join(save_path, f'Building Safety Release {this_month}.docx')
    DR.save(output_path)



# Call the main function
if __name__ == "__main__":
    create_DR()

# End the timer and print the elapsed time
end_time = time.time()
elapsed_time = end_time - start_time
print(f"Program executed in {elapsed_time:.2f} seconds")
print('Done! Have a nice day :)')