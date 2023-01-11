# Input: xml file name, project name, processing operator email address, location of xml file
xml_name = "C:/Users/combi/bin/NAV_xmls/1859_6732.xml"
site_name = 'WHT Q1 SQ'
assignee_email = 'lorenzo.combi@tre-altamira.com'
#xml_folder_path = "C:/Users/combi/bin/NAV_xmls/"
csv_folder_path = "C:/Users/combi/bin/Asana_csvs/"

import xml.dom.minidom
import pandas as pd
from datetime import datetime
import os

# dates in NAV changed format at some point: try different format to read dates from NAV xml
def get_date(date_str):
    date_formats = ["%d/%m/%y", "%y-%m-%d", "%Y-%m-%d","%Y/%m/%d","%m/%d/%y"]
    for date_format in date_formats:
        try:
            return datetime.strptime( date_str, date_format )
        except:
            pass

# Parse xml
doc = xml.dom.minidom.parse(xml_name)
# Get some attributes: Job Order number and delivery date (to the client), expected processing init and end
JOnumber = doc.getElementsByTagName("JobOrderCode")[0].firstChild.nodeValue
delivery_code = doc.getElementsByTagName("Delivery")[0].getElementsByTagName("Code")[0].firstChild.nodeValue
delivery_date = doc.getElementsByTagName("Delivery")[0].getElementsByTagName("Date")[0].firstChild.nodeValue
# All dates are strings, change to datetime format
delivery_date = get_date(delivery_date[:10])
expected_init = get_date(doc.getElementsByTagName("DeliveryExpectedInit")[0].firstChild.nodeValue)
expected_end = get_date(doc.getElementsByTagName("DeliveryExpectedProcessingEnd")[0].firstChild.nodeValue)
# Check dates
print('Client delivery date is ' + delivery_date.strftime("%Y-%m-%d"))
correct_delivery_date = input('Is the client delivery date correct? yes/no: ')
if 'no' in correct_delivery_date:
    delivery_date = get_date(input('Insert correct delivery date (YYYY/MM/DD): '))
print('Expected init is ' + expected_init.strftime("%Y-%m-%d"))
correct_expected_init = input('Is the expected init correct? yes/no: ')
if 'no' in correct_expected_init:
    expected_init = get_date(input('Insert correct expected init (YYYY/MM/DD): '))
print('Expected end is ' + expected_end.strftime("%Y-%m-%d"))
correct_expected_end = input('Is the expected end correct? yes/no: ')
if 'no' in correct_expected_end:
    expected_end = get_date(input('Insert correct expected end (YYYY/MM/DD): '))


print(expected_init)
# Build  strings for task names: processing beta and finalization tasks
processing_task_name = 'JO' + JOnumber + '-' + delivery_code[0:2] + ' ' + site_name + ' processing beta'
finalization_task_name = 'JO' + JOnumber + '-' + delivery_code[0:2] + ' ' + site_name + ' finalization'
# Define standard checklist subtasks
checklist_beta = ["Kick-off meeting", "Dataset check", "DEM check", "QC merge", "QC differentials", \
    "makegrid parameters", "QC gammaima", "get cluster", "QC aps: k, planes, unwrap", "QC phasexp", \
            "Matlab: compare w/ previous deliveries", "Matlab: check density and coverage"]
checklist_finalization = ["Matlab: check output fields", "Check pdf"]
subtask_number = len(checklist_beta)+len(checklist_finalization)
# Dictionary with all info to create the Asana tasks
asana_proc_subtasks = {'Name': [processing_task_name, finalization_task_name] + checklist_beta + checklist_finalization, \
    'Assignee': [assignee_email, assignee_email] + [None]*subtask_number, \
        'Start Date': [expected_init, None] + [None]*subtask_number, \
            'Due Date': [expected_end, delivery_date]  + [None]*subtask_number, \
                'Dependencies': [None, processing_task_name] + [None]*subtask_number, \
                    'Processing Status': ['Waiting', 'Waiting'] + [None]*subtask_number, \
                        'Parent Task': [None, None] + [processing_task_name]*len(checklist_beta) + [finalization_task_name]*len(checklist_finalization), \
                            'Hours': [10, 1] + [0]*subtask_number}

pd.DataFrame.from_dict(asana_proc_subtasks).to_csv(csv_folder_path + 'Asana_' + os.path.basename(xml_name)[:-4] + '_' + site_name + '.csv', index = False)

# asana_proc_subtasks_df = pd.DataFrame.from_dict(asana_proc_subtasks)
# print(asana_proc_subtasks_df)

# Build dictionary with info for Asana tasks
# asana_proc = {'Name': [processing_task_name, finalization_task_name], \
#    'Assignee': [assignee_email, assignee_email], \
#        'Start Date': [expected_init.strftime("%Y-%m-%d"), None], \
#            'Due Date': [expected_end.strftime("%Y-%m-%d"), delivery_date.strftime("%Y-%m-%d")], \
#               'Dependencies': [None, processing_task_name], \
#                    'Status': ['Waiting', 'Waiting']}
# Write csv file from pandas DataFrame build from dictionary
# pd.DataFrame.from_dict(asana_proc).to_csv(csv_folder_path + 'Asana_' + xml_name[:-4] + '_' + site_name + '.csv', index = False)