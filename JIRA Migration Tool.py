from jira import JIRA
from atlassian import jira
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill
from sys import exit
from tkinter import *
from tkinter.filedialog import askopenfilename
import tkinter as tk
import os
import datetime
from time import sleep
import requests
from bs4 import BeautifulSoup
import json
import shutil
import concurrent.futures

# Migration Tool properties
current_version = '0.3'
config_file = 'config.json'

# JIRA Default configuration
JIRA_BASE_URL_OLD = ''
project_old = ''
JIRA_BASE_URL_NEW = ''
project_new = ''
team_project_prefix = ''

# JIRA API configs
JIRA_sprint_api = '/rest/agile/1.0/sprint/'
JIRA_team_api = '/rest/teams-api/1.0/team'
headers = {"Content-type": "application/json", "Accept": "application/json"}

# Sprints configs
old_sprints = {}
new_sprints = {}
old_board_id = 0
default_board_name = 'Shared Sprints'

# Excel configs
header_font = Font(color='00000000', bold=True)
header_fill = PatternFill(fill_type="solid", fgColor="8db5e2")
hyperlink = Font(underline='single', color='0563C1')
zoom_scale = 100
wb = Workbook()
default_validation = {}
excel_columns_validation_ranges = {'0': 'A2:A1048576',
                                   '1': 'B2:B1048576',
                                   '2': 'C2:C1048576',
                                   '3': 'D2:D1048576',
                                   '4': 'E2:E1048576',
                                   '5': 'F2:F1048576',
                                   '6': 'G2:G1048576',
                                   '7': 'H2:H1048576',
                                   '8': 'I2:I1048576',
                                   '9': 'J2:J1048576',
                                   '10': 'K2:K1048576',
                                   '11': 'L2:L1048576',
                                   '12': 'M2:M1048576',
                                   '13': 'N2:N1048576',
                                   '14': 'O2:O1048576',
                                   '15': 'P2:P1048576',
                                   '16': 'Q2:Q1048576',
                                   '17': 'R2:R1048576',
                                   '18': 'S2:S1048576',
                                   '19': 'T2:T1048576',
                                   '20': 'U2:U1048576',
                                   '21': 'V2:V1048576',
                                   '22': 'W2:W1048576',
                                   '23': 'X2:X1048576',
                                   '24': 'Y2:Y1048576',
                                   '25': 'Z2:Z1048576',
                                   '26': 'AA2:AA1048576',
                                   '27': 'AB2:AB1048576',
                                   '28': 'AC2:AC1048576',
                                   '29': 'AD2:AD1048576',
                                   '30': 'AE2:AE1048576',
                                   '31': 'AF2:AF1048576',
                                   '32': 'AG2:AG1048576',
                                   }

# Migration configs
temp_dir_name = 'Attachments_Temp/'
mapping_file = ''
jira_system_fields = ['Sprint', 'Epic Link', 'Epic Name', 'Story Points', 'Parent Link']
additional_mapping_fields = ['Description', 'Labels']
limit_migration_data = 0  # 0 if all
total_issues = 0
start_jira_key = 1
create_remote_link_for_old_issue = 0
username, password = ('', '')
auth = (username, password)
items_lst = {}
sub_tasks = {}
teams = {}
last_updated_date = 'YYYY-MM-DD'
updated_issues_num = 0
threads = 1
migrated_text = 'Migrated to'
verbose_logging = 0
migrate_fixversions_check = 1
migrate_components_check = 1
migrate_sprints_check = 1
migrate_attachments_check = 1
migrate_comments_check = 1
migrate_statuses_check = 1
migrate_links_check = 1
migrate_teams_check = 1
migrate_metadata_check = 1
validation_error = 0
delete_dummy_flag = 0
skip_migrated_flag = 1

# Mappings
issuetypes_mappings = {}
fields_mappings = {}
status_mappings = {}
field_value_mappings = {}

# Transitions mapping - for status changes
old_transitions = {}
new_transitions = {}


# Functions list
def read_excel(file_path=mapping_file, columns=0, rows=0, start_row=2):
    global issuetypes_mappings, fields_mappings, status_mappings, field_value_mappings, verbose_logging
    global JIRA_BASE_URL_OLD, JIRA_BASE_URL_NEW, project_old, project_new
    print("[START] Mapping file is opened for processing.")
    
    def remove_spaces(mapping):
        for lev_1, values in mapping.items():
            for lev_2, details in values.items():
                details_lst = []
                for data in details:
                    details_lst.append(data.strip())
                mapping[lev_1][lev_2] = details_lst
        return mapping
    
    mapping_type = 1
    
    try:
        df = load_workbook(file_path, data_only=True)
        excel_sheet_names = df.get_sheet_names()
        for excel_sheet_name in excel_sheet_names:
            value_mappings = {}
            df1 = df.get_sheet_by_name(excel_sheet_name)
            row_count = rows
            col_count = columns
            if rows == 0:
                row_count = df1.max_row
            if columns == 0:
                col_count = df1.max_column
            empty_row = ['' for i in range(col_count)]
            for row in df1.iter_rows(min_row=start_row, max_row=row_count, max_col=col_count):
                d = []
                for v in row:
                    val = v.value
                    if val is None:
                        val = ""
                    else:
                        val = str(val)
                    d.append(val)
                if set(d) != set(empty_row):
                    if excel_sheet_name == 'Project':
                        JIRA_BASE_URL_OLD = d[0]
                        JIRA_BASE_URL_NEW = d[2]
                        project_old = d[1]
                        project_new = d[3]
                        if d[4] == 'Target -> Source':
                            mapping_type = 0
                        break
                    elif excel_sheet_name == 'Issuetypes':
                        if mapping_type == 0:
                            issuetypes_mappings[d[2].strip()] = {"hierarchy": d[1].strip(), "issuetypes": d[3].split(',')}
                        else:
                            if d[1].strip() in issuetypes_mappings.keys():
                                issuetypes_mappings[d[1].strip()]["issuetypes"].append(d[0].strip())
                            else:
                                issuetypes_mappings[d[1].strip()] = {"hierarchy": '2', "issuetypes": [d[0].strip()]}
                    elif excel_sheet_name == 'Statuses':
                        if mapping_type == 0:
                            for issuetype in d[0].split(','):
                                if issuetype.strip() not in status_mappings.keys():
                                    status_mappings[issuetype.strip()] = {d[1].strip(): d[2].split(',')}
                                else:
                                    status_mappings[issuetype.strip()][d[1].strip()] = d[2].split(',')
                        else:
                            if d[0].strip() in status_mappings.keys():
                                if d[2].strip() in status_mappings[d[0].strip()].keys():
                                    status_mappings[d[0].strip()][d[2].strip()].append(d[1].strip())
                                else:
                                    status_mappings[d[0].strip()][d[2].strip()] = [d[1].strip()]
                            else:
                                status_mappings[d[0].strip()] = {d[2].strip(): [d[1].strip()]}
                            if d[2] == '' and verbose_logging == 1:
                                print("[WARNING] The mapping of '{}' status for '{}' Issuetype not found. Default status would be used.".format(d[1], d[0]))
                    elif excel_sheet_name == 'Fields':
                        if mapping_type == 0:
                            for issuetype in d[0].split(','):
                                if issuetype.strip() not in fields_mappings.keys():
                                    fields_mappings[issuetype.strip()] = {d[1].strip(): d[2].split(',')}
                                else:
                                    fields_mappings[issuetype.strip()][d[1].strip()] = d[2].split(',')
                        else:
                            if d[0].strip() in fields_mappings.keys():
                                if d[2].strip() in fields_mappings[d[0].strip()].keys():
                                    fields_mappings[d[0].strip()][d[2].strip()].append(d[1].strip())
                                else:
                                    fields_mappings[d[0].strip()][d[2].strip()] = [d[1].strip()]
                            else:
                                fields_mappings[d[0].strip()] = {d[2].strip(): [d[1].strip()]}
                            if d[2] == '' and verbose_logging == 1:
                                print("[WARNING] The mapping of '{}' field for '{}' Issuetype not found. Field values will be dropped.".format(d[1], d[0]))
                    else:
                        if len(d) <= 2:
                            try:
                                if mapping_type == 0:
                                    value_mappings[d[0].strip()] = d[1].split(',')
                                else:
                                    if d[1].strip() not in value_mappings.keys():
                                        value_mappings[d[1].strip()] = d[0].strip().split(',')
                                    else:
                                        value_mappings[d[1].strip()].extend(d[0].strip().split(','))
                            except:
                                print("[ERROR] Data on the sheet '{}' is invalid, Skipping...".format(excel_sheet_name))
                                continue
                        else:
                            try:
                                if mapping_type == 1:
                                    if d[1] + ' --> ' + d[2] not in value_mappings.keys():
                                        value_mappings[d[1] + ' --> ' + d[2]] = d[0].strip().split(',')
                                    else:
                                        value_mappings[d[1] + ' --> ' + d[2]].extend(d[0].strip().split(','))
                            except:
                                print("[ERROR] Data on the sheet '{}' is invalid, Skipping...".format(excel_sheet_name))
                                continue
            
            if excel_sheet_name not in ['Project', 'Issuetypes', 'Statuses', 'Fields']:
                field_value_mappings[excel_sheet_name] = value_mappings
    except:
        print("[ERROR] '{}' file not found. Mappings can't be processed.".format(file_path))
        os.system("pause")
        exit()
    for k, v in issuetypes_mappings.items():
        issues = []
        for issuetype in v['issuetypes']:
            issues.append(issuetype.strip())
        issuetypes_mappings[k]['issuetypes'] = issues
    
    status_mappings = remove_spaces(status_mappings)
    fields_mappings = remove_spaces(fields_mappings)
    field_value_mappings = remove_spaces(field_value_mappings)
    print("[END] Mapping data has been successfully processed.")


def get_transitions(project, jira_url, new=False):
    global old_transitions, new_transitions, auth, migrate_statuses_check, headers
    print("[START] Retrieving Transitions and Statuses for '{}' project from JIRA.".format(project))
    
    def get_workflows(project, jira_url, new):
        global sub_tasks, auth
        url = jira_url + '/rest/projectconfig/1/workflowscheme/' + project
        r = requests.get(url, auth=auth, headers=headers)
        workflow_schema_string = r.content.decode('utf-8')
        workflow_schema_details = json.loads(workflow_schema_string)
        workflows = {}
        issuetypes = {}
        for issuetype in workflow_schema_details['issueTypes']:
            issuetypes[issuetype['id']] = issuetype['name']
            if new is True and issuetype['subTask'] is True:
                sub_tasks[issuetype['name']] = issuetype['id']
        for workflow in workflow_schema_details['mappings']:
            workflows[workflow['name']] = [issuetypes[i] for i in workflow['issueTypes']]
        return workflows
    
    try:
        transitions = {}
        for workflow_name, workflow_details in get_workflows(project, jira_url, new).items():
            for issuetype in workflow_details:
                url0 = jira_url + '//rest/projectconfig/1/workflow?workflowName=' + workflow_name + '&projectKey=' + project
                url1 = jira_url + '/rest/projectconfig/1/workflow?workflowName=' + workflow_name + '&projectKey=' + project
                r = requests.get(url0, auth=auth, headers=headers)
                if r.status_code == 200:
                    workflow_string = r.content.decode('utf-8')
                else:
                    r = requests.get(url1, auth=auth, headers=headers)
                    workflow_string = r.content.decode('utf-8')
                workflow_data = json.loads(workflow_string)
                transition_details = []
                for status in workflow_data["sources"]:
                    for target in status["targets"]:
                        transition_details.append([status["fromStatus"]["name"], target['transitionName'], target['toStatus']['name']])
                transitions[issuetype] = transition_details
        if new is False:
            old_transitions = transitions
        else:
            new_transitions = transitions
        print("[END] Transitions and Statuses for '{}' project has been successfully retrieved.".format(project))
    except Exception as e:
        migrate_statuses_check = 0
        print("[WARNING] No PROJECT ADMIN right for the '{}' project. Statuses WILL NOT be updated / migrated.".format(project))
        print("[ERROR] Transitions and Statuses can't be retrieved due to '{}'".format(e))


def get_hierarchy_config():
    global sub_tasks, issuetypes_mappings, issue_details_new
    
    for issuetype, details in issuetypes_mappings.items():
        try:
            if issuetype in sub_tasks.keys():
                issuetypes_mappings[issuetype]['hierarchy'] = '3'
            elif 'Epic Link' in issue_details_new[issuetype].keys():
                issuetypes_mappings[issuetype]['hierarchy'] = '2'
            elif 'Epic Name' in issue_details_new[issuetype].keys():
                issuetypes_mappings[issuetype]['hierarchy'] = '1'
            else:
                issuetypes_mappings[issuetype]['hierarchy'] = '0'
        except:
            print("[ERROR] '{}' Issue Type not available in mapping file.".format(issuetype))


def prepare_template_data():
    global old_transitions, new_transitions, issue_details_old, default_validation, jira_system_fields
    global JIRA_BASE_URL_OLD, project_old, JIRA_BASE_URL_NEW, project_new, migrate_statuses_check
    global additional_mapping_fields
    
    template_excel = {}
    old_statuses, new_statuses, old_issuetypes, new_issuetypes = ([], [], [], [])
    
    def calculate_statuses(transitions):
        issuetypes_lst = []
        issuetype_statuses = {}
        for k, v in transitions.items():
            statuses_lst = []
            issuetypes_lst.append(k)
            for l in v:
                statuses_lst.append(l[0])
                statuses_lst.append(l[2])
            issuetype_statuses[k] = list(set(statuses_lst))
        statuses_lst = []
        for k, v in issuetype_statuses.items():
            for status in v:
                statuses_lst.append([k, status, ''])
        return statuses_lst, list(set(issuetypes_lst))
    
    # Project details
    project_details = [['Source Project JIRA URL', 'Source Project JIRA Key', 'Target Project JIRA URL', 'Target Project JIRA Key', 'Template type']]
    project_details.append([JIRA_BASE_URL_OLD, project_old, JIRA_BASE_URL_NEW, project_new, 'Source -> Target'])
    
    # IssueTypes
    if migrate_statuses_check == 1:
        old_statuses, old_issuetypes = calculate_statuses(old_transitions)
        new_statuses, new_issuetypes = calculate_statuses(new_transitions)
    else:
        for issuetype in issue_details_old.keys():
            old_issuetypes.append(issuetype)
        old_issuetypes = list(set(old_issuetypes))
        for issuetype in issue_details_new.keys():
            new_issuetypes.append(issuetype)
        new_issuetypes = list(set(new_issuetypes))
    issue_types_map_lst = [['Source Issue type', 'Target Issue Type']]
    for o_it in old_issuetypes:
        issue_types_map_lst.append([o_it, ''])
        
    # Fields
    fields_map_lst = [['Source Issue Type', ' Source Field Name', 'Target Field Name']]
    for issuetype, fields in issue_details_old.items():
        for field, details in fields.items():
            if details['custom'] is True and field not in jira_system_fields:
                fields_map_lst.append([issuetype, field, ''])
    
    new_fields_val = additional_mapping_fields[:]
    for issuetype, fields in issue_details_new.items():
        for field, details in fields.items():
            if details['custom'] is True and field not in jira_system_fields:
                new_fields_val.append(field.title())
    
    # Statuses
    if migrate_statuses_check == 1:
        statuses_map_lst = [['Source Issue Type', 'Source Status', 'Target Status']]
        statuses_map_lst.extend(old_statuses)
    
    # Priority
    priority_map_lst = [['Source Priority', 'Target Priority']]
    priority_old_lst = []
    priority_new_lst = []
    for field_values in issue_details_old.values():
        if 'Priority' in field_values.keys():
            for p in field_values['Priority']['allowed values']:
                priority_old_lst.append(p)
    priority_old_lst = list(set(priority_old_lst))
    for priority in priority_old_lst:
        priority_map_lst.append([priority, ''])
    for field_values in issue_details_new.values():
        if 'Priority' in field_values.keys():
            for p in field_values['Priority']['allowed values']:
                priority_new_lst.append(p)
    
    # Combine all data under one dictionary
    template_excel['Project'] = project_details
    template_excel['Issuetypes'] = issue_types_map_lst
    template_excel['Fields'] = fields_map_lst
    if migrate_statuses_check == 1:
        template_excel['Statuses'] = statuses_map_lst
    template_excel['Priority'] = priority_map_lst
    
    # Other fields
    for field_values in issue_details_new.values():
        for field_name, field_data in field_values.items():
            if field_name not in template_excel.keys() and field_name not in jira_system_fields:
                if field_data['type'] != 'option-with-child':
                    new_field_map_lst = [["Source '{}'".format(field_name), "Target '{}'".format(field_name)]]
                else:
                    new_field_map_lst = [["Source '{}'".format(field_name), "Target Level 1 of '{}'".format(field_name), "Target Level 2 of '{}'".format(field_name)]]
                if field_data['custom'] is True and field_data['allowed values'] is not None:
                    for value in field_data['allowed values']:
                        if field_data['type'] != 'option-with-child':
                            new_field_map_lst.append(['', value])
                        else:
                            new_field_map_lst.append(['', value[0], value[1]])
                    
                    template_excel[field_name] = new_field_map_lst
    
    # Calculate and update validation
    if migrate_statuses_check == 1:
        new_statuses_val = []
        for i in new_statuses:
            new_statuses_val.append(i[1].title())
        default_validation['Statuses'] = '"' + get_str_from_lst(list(set(new_statuses_val)), spacing='') + '"'
    default_validation['Fields'] = '"' + get_str_from_lst(list(set(new_fields_val)), spacing='') + '"'
    new_issuetypes_val = []
    for i in new_issuetypes:
        try:
            new_issuetypes_val.append(i.title())
        except:
            new_issuetypes_val.append(i)
    default_validation['Issuetypes'] = '"' + get_str_from_lst(list(set(new_issuetypes_val)), spacing='') + '"'
    default_validation['Priority'] = '"' + get_str_from_lst(list(set(priority_new_lst)), spacing='') + '"'
    
    return template_excel


def get_issues_by_jql(jira, jql, types=None, sprint=None, details=None, max_result=limit_migration_data, issue_details=None):
    """This function returns list of JIRA keys for provided list of JIRA JQL queries"""
    global old_sprints, items_lst, limit_migration_data, total_issues
    
    print("[START] The list of all Issues are retrieving from JIRA.")
    auth_jira = jira
    issues, items = ([], [])
    start_idx, block_num, block_size = (0, 0, 100)
    if sprint is not None and issue_details is not None:
        sprint_field_id = issue_details['Story']['Sprint']['id']
    if max_result != 0 and block_size > max_result:
        block_size = max_result
    while True:
        if block_size > max_result:
            block_size = max_result
        start_idx = block_num * block_size
        try:
            tmp_issues = auth_jira.search_issues(jql_str=jql, startAt=start_idx, maxResults=block_size, fields='key, issuetype')
            if len(tmp_issues) == 0:
                # Retrieve issues until there are no more to come
                break
        except Exception as er:
            print("[ERROR] Exception while retrieving JIRA Data: '{}'".format(er.text))
            break
        issues.extend(tmp_issues)
        block_num += 1
        max_result -= block_size
        if max_result == 0:
            # If only first max_result items are required
            break
    print("[END] The list of all Issues has been successfully retrieved from JIRA.", '', sep='\n')
    if types is not None:
        type_lst = list(set([i.fields.issuetype.name for i in issues]))
        for issuetype in type_lst:
            if issuetype not in items_lst.keys():
                items_lst[issuetype] = set()
        for i in issues:
            issue = auth_jira.issue(i.key)
            items_lst[i.fields.issuetype.name].add(issue.key)
    elif sprint is not None or details is not None:
        type_lst = list(set([i.fields.issuetype.name for i in issues]))
        total_issues = len(issues)
        for issuetype in type_lst:
            if issuetype not in items_lst.keys():
                items_lst[issuetype] = set()
        print("[INFO] Sprint retrieval from project was started. It could take some time... Please wait...")
        for i in issues:
            issue = auth_jira.issue(i.key)
            items_lst[i.fields.issuetype.name].add(issue.key)
            if sprint is not None:
                issue_sprints = eval('issue.fields.' + sprint_field_id)
                if issue_sprints is not None:
                    for sprint in issue_sprints:
                        sprint_id, name, state, startDate, endDate = ('', '', '', '', '')
                        for attr in sprint[sprint.find('[')+1:-1].split(','):
                            if 'id=' in attr:
                                sprint_id = attr.split('id=')[1]
                            if 'name=' in attr:
                                name = attr.split('name=')[1]
                            if 'state=' in attr:
                                state = attr.split('state=')[1]
                            if 'startDate=' in attr:
                                startDate = '' if attr.split('startDate=')[1] == '<null>' else attr.split('startDate=')[1]
                            if 'endDate=' in attr:
                                endDate = '' if attr.split('endDate=')[1] == '<null>' else attr.split('endDate=')[1]
                        if name not in old_sprints.keys():
                            old_sprints[name] = {"id": sprint_id, "startDate": startDate, "endDate": endDate, "state": state.upper()}
    else:
        if len(issues) > 0:
            return list(set([i.key for i in issues]))
        else:
            return None


def get_str_from_lst(lst, sep=',', spacing=' '):
    """This function returns list as comma separated string - for exporting in excel"""
    if lst is None:
        return None
    elif type(lst) != list:
        return str(lst)
    st = ''
    for i in lst:
        if i != '':
            st += str(i).strip() + sep + spacing
    if spacing == ' ':
        st = st[0:-2]
    else:
        st = st[0:-1]
    return st


def create_temp_folder(folder):
    local_folder = folder
    if os.path.exists(local_folder):
        for filename in os.listdir(local_folder):
            file_path = os.path.join(local_folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print('Failed to delete %s. Reason: %s' % (file_path, e))
        print("[INFO] Folder '{}' has been cleaned up.".format(local_folder))
    else:
        os.mkdir(local_folder)
        print("[INFO] Folder '{}' has been created".format(local_folder))


def clean_temp_folder(folder):
    shutil.rmtree(folder)


def get_all_shared_teams():
    global teams, verbose_logging, auth, headers
    print("[START] Reading ALL available shared teams.")
    i = 1
    while True:
        url_retrieve = JIRA_BASE_URL_NEW + JIRA_team_api + '?size=100&page=' + str(i)
        r = requests.get(url=url_retrieve, auth=auth, headers=headers)
        teams_string = r.content.decode('utf-8')
        if len(teams_string) < 5:
            break
        teams_lst = json.loads(teams_string)
        for team in teams_lst:
            teams[team['title']] = team['id']
        if verbose_logging == 1:
            print('[INFO] Teams retrieved from JIRA so far: ', len(teams))
        i += 1
    print("[END] All teams has been loaded for further items processing.")


def get_team_id(team_name):
    global teams, team_project_prefix, auth
    
    def create_new_team():
        global teams, auth, headers
        url_create = JIRA_BASE_URL_NEW + JIRA_team_api
        team_name_to_create = team_project_prefix + team_name
        body = eval('{"title": team_name_to_create, "shareable": "true"}')
        r = requests.post(url_create, json=body, auth=auth, headers=headers)
        team_id = int(r.content.decode('utf-8'))
        teams[team_name_to_create] = team_id
        return str(team_id)
    
    team_name_to_check = team_project_prefix + team_name
    if team_name_to_check in teams.keys():
        return str(teams[team_name_to_check])
    return create_new_team()


def migrate_sprints(board_id=old_board_id, proj_old=None, project=project_new, name=default_board_name, param='FUTURE'):
    global old_sprints, new_sprints, jira_old, jira_new, limit_migration_data, limit_migration_data, auth
    global max_processing_key, start_jira_key, headers
    
    print()
    new_board, n = (0, 0)
    for board in jira_new.boards():
        if board.name == name:
            new_board = board.id
    if new_board == 0:
        new_board = jira_new.create_board(name, project, location_type='project')
    
    for n_sprint in jira_new.sprints(board_id=new_board):
        new_sprints[n_sprint.name] = {"id": n_sprint.id, "state": n_sprint.state}
    if param == 'FUTURE':
        if proj_old is None:
            print("[INFO] Sprints to be migrated from board '{}'.".format(board_id))
            for sprint in jira_old.sprints(board_id=board_id):
                if sprint.name not in new_sprints.keys():
                    try:
                        old_sprint = jira_old.sprint(sprint.id)
                        old_sprints[sprint.name] = {"id": sprint.id, "startDate": old_sprint.startDate, "endDate": old_sprint.endDate, "state": old_sprint.state.upper()}
                    except:
                        old_sprints[sprint.name] = {"id": sprint.id, "startDate": '', "endDate": '', "state": sprint.state.upper()}
                n += 1
                if (n % 20) == 0:
                    print("[INFO] Downloaded metadata for {} out of {} Sprints so far...".format(n, len(jira_old.sprints(board_id=board_id))))
        else:
            print("[INFO] All Sprints to be migrated from old '{}' project and will be added into new '{}' project, '{}' board.".format(proj_old, project, name))
            if limit_migration_data != 0:
                jql_sprints = 'project = {} AND key >= {} AND key <= {} order by key ASC'.format(project_old, start_jira_key, max_processing_key)
            else:
                jql_sprints = 'project = {} AND key >= {} order by key ASC'.format(proj_old, start_jira_key)
            get_issues_by_jql(jira_old, jql=jql_sprints, sprint=True, issue_details=issue_details_old)
        
        print("[START] Missing Sprints to be created...")
        for o_sprint_name, o_sprint_details in old_sprints.items():
            if o_sprint_name not in new_sprints.keys():
                try:
                    new_sprint = jira_new.create_sprint(name=o_sprint_name, board_id=name, startDate=old_sprints[o_sprint_name]['startDate'], endDate=old_sprints[o_sprint_name]['endDate'])
                    new_sprints[new_sprint.name] = {"id": new_sprint.id, "state": new_sprint.state}
                except:
                    print("[WARNING] Sprint '{}' can't be migrated. It has been deleted or access to board is restricted. Skipped...".format(o_sprint_name))
        print("[END] Sprints have been created with 'Future' states.")
    else:
        print("[START] Sprint statuses to be updated to '{}'.".format(param))
        for o_sprint_name, o_sprint_details in old_sprints.items():
            if o_sprint_name in new_sprints.keys():
                url = JIRA_BASE_URL_NEW + JIRA_sprint_api + str(new_sprints[o_sprint_name]["id"])
                if param == 'ACTIVE' and old_sprints[o_sprint_name]['endDate'] != new_sprints[o_sprint_name]["state"] and old_sprints[o_sprint_name]['endDate'] != 'FUTURE':
                    # Mark Sprint as Active
                    if jira_new.sprint(new_sprints[o_sprint_name]["id"]).state == 'FUTURE':
                        body = {"state": "ACTIVE"}
                        r = requests.post(url, json=body, auth=auth, headers=headers)
                    if param == 'CLOSED' and jira_new.sprint(new_sprints[o_sprint_name]["id"]).state == 'ACTIVE' and old_sprints[o_sprint_name]['endDate'] == 'CLOSED':
                        # Mark Sprint as Closed - should be Active beforehand
                        body = {"state": "CLOSED"}
                        r = requests.post(url, json=body, auth=auth, headers=headers)
        print("[END] Sprint statuses have been updated to '{}'.".format(param), '', sep='\n')


def migrate_components():
    print("[START] Components migration has been started.")
    old_components = jira_old.project_components(project_old)
    new_components = jira_new.project_components(project_new)
    
    new_components_lst = []
    for new_component in new_components:
        new_components_lst.append(new_component.name)
    
    for component in old_components:
        description, assignee_type, lead_name, assignee_valid = (None, None, None, None)
        if component.name not in new_components_lst and component.archived is False:
            if hasattr(component, 'description'):
                description = component.description
            if hasattr(component, 'assigneeType'):
                assignee_type = component.assigneeType
            if hasattr(component, 'lead') and hasattr(component.lead, 'name'):
                lead_name = component.lead.name
            if hasattr(component, 'isAssigneeTypeValid'):
                assignee_valid = component.isAssigneeTypeValid
            try:
                jira_new.create_component(component.name, project_new, description=description, leadUserName=lead_name, assigneeType=assignee_type, isAssigneeTypeValid=assignee_valid)
            except Exception as e:
                print('Exception: {}'.format(e.text))
    print("[END] All components have been succsessfully migrated.", '', sep='\n')


def migrate_versions():
    print("[START] FixVersions (Releases) migration has been started.")
    old_versions = jira_old.project_versions(project_old)
    new_versions = jira_new.project_versions(project_new)
    
    new_versions_lst = []
    for new_version in new_versions:
        new_versions_lst.append(new_version.name)
    
    for version in old_versions:
        description, release_date, start_date, archieved, released = (None, None, None, None, None)
        if version.name not in new_versions_lst:
            if hasattr(version, 'description'):
                description = version.description
            if hasattr(version, 'releaseDate'):
                release_date = version.releaseDate
            if hasattr(version, 'startDate'):
                start_date = version.startDate
            if hasattr(version, 'archieved'):
                archieved = version.archieved
            if hasattr(version, 'released'):
                released = version.released
            try:
                jira_new.create_version(version.name, project_new, description=description, releaseDate=release_date, startDate=start_date, archived=archieved, released=released)
            except Exception as e:
                print('Exception: {}'.format(e.text))
    print("[END] All FixVersions (Releases) have been succsessfully migrated.", '', sep='\n')


def migrate_comments(old_issue, new_issue):
    for comment in jira_old.comments(old_issue):
        comment_match = 0
        new_data = eval("'*[' + comment.author.displayName + '|~' + comment.author.name + ']* added on *_' + comment.created[:10] + ' ' + comment.created[11:19] + '_*: \\\\\\ '")
        len_new_data = len(new_data)
        for new_comment in jira_new.comments(new_issue):
            if comment.body == new_comment.body[len_new_data:] or comment.body == new_comment.body:
                comment_match = 1
        if comment_match == 0:
            data = eval("new_data + comment.body")
            jira_new.add_comment(new_issue, body=str(data))


def migrate_links(old_issue, new_issue):
    outward_issue_links = {}
    inward_issue_links = {}
    try:
        old_links = old_issue.fields.issuelinks
    except:
        old_links = []
    try:
        new_links = new_issue.fields.issuelinks
    except:
        new_links = []
    
    for link in new_links:
        if hasattr(link, "outwardIssue"):
            outward_issue_links[link.outwardIssue.key] = link.type.outward
        if hasattr(link, "inwardIssue"):
            inward_issue_links[link.inwardIssue.key] = link.type.inward
    for link in old_links:
        if hasattr(link, "outwardIssue"):
            new_id = link.outwardIssue.key.replace(project_old, project_new)
            if new_id not in outward_issue_links.keys() or (new_id in outward_issue_links.keys()
                                                            and link.type.outward != outward_issue_links[new_id]):
                try:
                    jira_new.create_issue_link(link.type.outward, new_issue.key, new_id)
                except:
                    pass
        if hasattr(link, "inwardIssue"):
            new_id = link.inwardIssue.key.replace(project_old, project_new)
            if new_id not in inward_issue_links.keys() or (new_id in inward_issue_links.keys()
                                                           and link.type.inward != inward_issue_links[new_id]):
                try:
                    jira_new.create_issue_link(link.type.inward, new_issue.key, new_id)
                except:
                    pass

def migrate_attachments(old_issue, new_issue):
    global temp_dir_name
    new_attachments = []
    if new_issue.fields.attachment:
        for new_attachment in new_issue.fields.attachment:
            new_attachments.append(new_attachment.filename)
    if old_issue.fields.attachment:
        for attachment in old_issue.fields.attachment:
            if attachment.filename not in new_attachments:
                file = attachment.get()
                filename = attachment.filename
                full_name = os.path.join(temp_dir_name, filename)
                with open(full_name, 'wb') as f:
                    f.write(file)
                with open(full_name, 'rb') as file_new:
                    jira_new.add_attachment(new_issue.key, file_new, filename)
                if os.path.exists(full_name):
                    os.remove(full_name)


def migrate_status(new_issue, old_issue):
    global new_transitions
    
    def get_new_status(old_status, issue_type):
        global status_mappings
        for n_status, o_statuses in status_mappings[issue_type].items():
            for o_status in o_statuses:
                if old_status.upper() == o_status.upper():
                    return n_status
        return None
    
    def find_shortest_path(graph, start, end, path):
        path = path + [start]
        if start == end:
            return path
        if start not in graph.keys():
            return None
        shortest = None
        for node in graph[start]:
            if node not in path:
                new_path = find_shortest_path(graph, node, end, path)
                if new_path and (not shortest or len(new_path) < len(shortest)):
                    shortest = new_path
        return shortest
    
    resolution = None
    issue_type = old_issue.fields.issuetype.name
    new_status = get_new_status(old_issue.fields.status.name, issue_type)
    old_status = new_issue.fields.status.name
    if old_issue.fields.resolution:
        resolution = old_issue.fields.resolution.name
    
    if new_status == '':
        return
    graph = {}
    for issuetype, details in issuetypes_mappings.items():
        if issue_type in details['issuetypes']:
            new_issue_type = issuetype
            break
    
    for k, v in new_transitions.items():
        if k == new_issue_type:
            for t in v:
                if t[0].upper() in graph.keys():
                    graph[t[0].upper()].append(t[2].upper())
                else:
                    graph[t[0].upper()] = [t[2].upper()]
    if old_status is None:
        old_status = graph[new_issue_type][0][0]
    transition_path = find_shortest_path(graph, old_status.upper(), new_status.upper(), [])
    if transition_path is None:
        return
    
    status_transitions = []
    for k, v in new_transitions.items():
        if k == new_issue_type:
            for i in range(1, len(transition_path)):
                for t in v:
                    if t[0].upper() == transition_path[i-1] and t[2].upper() == transition_path[i]:
                        status_transitions.append(t[1])
    
    for s in status_transitions:
        if resolution is None:
            jira_new.transition_issue(new_issue, transition=s)
        else:
            try:
                jira_new.transition_issue(new_issue, transition=s, fields={"resolution": {"name": resolution}})
            except:
                jira_new.transition_issue(new_issue, transition=s)


def migrate_issues(issuetype):
    global items_lst, threads
    
    def process_issue(key):
        global items_lst, jira_new, project_new, jira_old, migrate_comments_check, migrate_links_check, migrated_text
        global migrate_attachments_check, migrate_statuses_check, migrate_metadata_check, create_remote_link_for_old_issue

        new_issue_type = ''
        new_issue_key = project_new + '-' + key.split('-')[1]
        try:
            old_issue = jira_old.issue(key)
        except:
            return
        try:
            new_issue = jira_new.issue(new_issue_key)
        except:
            new_issue = create_dummy_issue(jira_new, project_new, get_minfields_issuetype(issue_details_new)[0], get_minfields_issuetype(issue_details_new)[1], old_issue)
        if migrate_metadata_check == 1:
            issue_type = old_issue.fields.issuetype.name
            for issuetype, details in issuetypes_mappings.items():
                if issue_type in details['issuetypes']:
                    new_issue_type = issuetype
                    break
            update_new_issue_type(old_issue, new_issue, new_issue_type)
        if migrate_comments_check == 1:
            migrate_comments(old_issue, new_issue)
        if migrate_links_check == 1:
            migrate_links(old_issue, new_issue)
        if migrate_attachments_check == 1:
            migrate_attachments(old_issue, new_issue)
        if migrate_statuses_check == 1:
            migrate_status(new_issue, old_issue)
        if create_remote_link_for_old_issue == 1:
            remote_link_exist = 0
            try:
                for r_link in jira_old.remote_links(old_issue.key):
                    if r_link.object.title == new_issue.key and r_link.relationship == migrated_text:
                        remote_link_exist = 1
            except:
                pass
            if remote_link_exist == 0:
                atlassian_jira_old.create_or_update_issue_remote_links(old_issue.key, JIRA_BASE_URL_NEW + '/browse/' + new_issue.key, title=new_issue.key, relationship='Migrated to')

    for type in issuetypes_mappings[issuetype]['issuetypes']:
        if type in items_lst.keys():
            print()
            print("[START] Copying from old '{}' Issuetype to new '{}' Issuetype...".format(type, issuetype))
            with concurrent.futures.ThreadPoolExecutor(max_workers=threads) as executor:
                futures = [executor.submit(process_issue, key) for key in items_lst[type]]
        else:
            print("[INFO] No issues under '{}' issuetype were found. Skipping...".format(type))
            continue
        print("[END] '{}' issuetype has been migrated to '{}' Issuetype.".format(type, issuetype), '', sep='\n')


def get_fields_list_by_project(jira, project):
    auth_jira = jira
    allfields = auth_jira.fields()
    
    def retrieve_custom_field(field_id):
        for field in allfields:
            if field['id'] == field_id:
                return field['custom']
    
    proj = auth_jira.project(project)
    try:
        project_fields = auth_jira.createmeta(projectKeys=proj, expand='projects.issuetypes.fields')
    except:
        print("[ERROR] NO ACCESS to the '{}' project.".format(proj))
        return {}
    is_types = project_fields['projects'][0]['issuetypes']
    issuetype_fields = {}
    for issuetype in is_types:
        issuetype_name = issuetype['name']
        issuetype_fields[issuetype_name] = {}
        for field_id in issuetype['fields']:
            field_name = issuetype['fields'][field_id]['name']
            allowed_values = []
            if 'allowedValues' in issuetype['fields'][field_id]:
                for i in issuetype['fields'][field_id]['allowedValues']:
                    if 'children' in i:
                        for ch in i['children']:
                            if 'value' in ch.keys():
                                allowed_values.append([i['value'], ch['value']])
                    elif 'name' in i:
                        allowed_values.append(i['name'])
                    elif 'value' in i:
                        allowed_values.append(i['value'])
            field_attributes = {'id': field_id, 'required': issuetype['fields'][field_id]['required'],
                                'custom': retrieve_custom_field(field_id),
                                'type': issuetype['fields'][field_id]['schema']['type'],
                                'custom type': None if 'custom' not in issuetype['fields'][field_id]['schema'] else issuetype['fields'][field_id]['schema']['custom'].replace('com.atlassian.jira.plugin.system.customfieldtypes:', ''),
                                'allowed values': None if allowed_values == [] else allowed_values,
                                'default value': None if issuetype['fields'][field_id]['hasDefaultValue'] is False else issuetype['fields'][field_id]['defaultValue']['name'] if 'name' in issuetype['fields'][field_id]['defaultValue'] else issuetype['fields'][field_id]['defaultValue']['value'],
                                'validated': True if 'allowedValues' in issuetype['fields'][field_id] else False}
            issuetype_fields[issuetype_name][field_name] = field_attributes
    return issuetype_fields


def load_file():
    global mapping_file, issues, header
    dir_name = os.getcwd()
    mapping_file = askopenfilename(initialdir=dir_name, title="Select file", filetypes=(("Migration JIRA Template", ".xlsx .xls"), ("all files", "*.*")))
    file.delete(0, END)
    file.insert(0, mapping_file)


def create_excel_sheet(sheet_data, title):
    global JIRA_BASE_URL, header, output_excel, default_validation, issue_details_new, issue_details_old
    global jira_system_fields, additional_mapping_fields
    wb.create_sheet(title)
    ws = wb.get_sheet_by_name(title)
    
    start_column = 1
    start_row = 1
    
    # Creating Excel sheet based on data
    for i in range(len(sheet_data)):
        for y in range(len(sheet_data[i])):
            try:
                ws.cell(row=start_row, column=start_column+y).value = sheet_data[i][y]
            except:
                converted_value = ''
                for letter in sheet_data[i][y]:
                    if letter.isalpha() or letter.isnumeric() or letter in [',', '.', ';', ':', '&', '"', "'", ' ']:
                        converted_value += letter
                    else:
                        converted_value += '?'
                ws.cell(row=start_row, column=start_column+y).value = converted_value
        start_row += 1
    
    for y in range(1, ws.max_column+1):
        ws.cell(row=1, column=y).fill = header_fill
        ws.cell(row=1, column=y).font = header_font
    
    # Column width formatting
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        if length > 80:
            ws.column_dimensions[column_cells[0].column_letter].width = 80
        else:
            ws.column_dimensions[column_cells[0].column_letter].width = length + 4
    
    # Data Validation added - for short lists in the beginning
    if title == 'Project':
        fields_val = {}
        for issuetype, fields in issue_details_new.items():
            fields_val[issuetype] = additional_mapping_fields[:]
            for field in fields.keys():
                if issue_details_new[issuetype][field]['custom'] is True and field not in jira_system_fields:
                    fields_val[issuetype].append(field)
        
        for issuetype, fields in fields_val.items():
            issuetype_fields = []
            for field in fields:
                issuetype_fields.append(field)
            start_row = 2
            start_column = ws.max_column + 1
            
            for f in issuetype_fields:
                ws.cell(row=start_row+1, column=start_column).value = f
                start_row += 1
            col_letter = ws.cell(row=start_row, column=start_column).column_letter
            issuetype_value = DefinedName(issuetype.replace('-', '_').replace(' ', '__'), attr_text='Project!$' + col_letter + '$3:$' + col_letter + '$' + str(len(issuetype_fields) + 2))
            wb.defined_names.append(issuetype_value)
            ws.column_dimensions[col_letter].hidden = True
        
        statuses_val = {}
        for issuetype, statuses in new_transitions.items():
            statuses_lst = []
            for status in statuses:
                statuses_lst.append(status[0])
                statuses_lst.append(status[2])
            statuses_lst = list(set(statuses_lst))
            statuses_val[issuetype] = statuses_lst
        
        for issuetype, statuses in statuses_val.items():
            issuetype_statuses = []
            for status in statuses:
                issuetype_statuses.append(status)
            start_row = 2
            start_column = ws.max_column + 1
            
            for s in issuetype_statuses:
                ws.cell(row=start_row+1, column=start_column).value = s
                start_row += 1
            col_letter = ws.cell(row=start_row, column=start_column).column_letter
            issuetype_value = DefinedName(issuetype.replace('-', '_').replace(' ', '__') + 'STATUS', attr_text='Project!$' + col_letter + '$3:$' + col_letter + '$' + str(len(issuetype_statuses) + 2))
            wb.defined_names.append(issuetype_value)
            ws.column_dimensions[col_letter].hidden = True
    
    if title == 'Issuetypes':
        issuetypes_val = DataValidation(type="list", formula1=default_validation['Issuetypes'], allow_blank=False)
        ws.add_data_validation(issuetypes_val)
        issuetypes_val.add(excel_columns_validation_ranges['1'])
    
    if title == 'Fields':
        issuetypes_val = DataValidation(type="list", formula1='INDIRECT(SUBSTITUTE(SUBSTITUTE(VLOOKUP($A2,Issuetypes!$A$2:$B$' + str(len(issue_details_old.keys()) + 1) + ',2,FALSE)," ","__"),"-","_"))', allow_blank=False)
        ws.add_data_validation(issuetypes_val)
        issuetypes_val.add(excel_columns_validation_ranges['2'])
    
    if title == 'Statuses':
        issuetypes_val = DataValidation(type="list", formula1='INDIRECT(CONCATENATE(SUBSTITUTE(SUBSTITUTE(VLOOKUP($A2,Issuetypes!$A$2:$B$' + str(len(issue_details_old.keys()) + 1) + ',2,FALSE)," ","__"),"-","_"),"STATUS"))', allow_blank=False)
        ws.add_data_validation(issuetypes_val)
        issuetypes_val.add(excel_columns_validation_ranges['2'])
    
    if title == 'Priority':
        priority_val = DataValidation(type="list", formula1=default_validation['Priority'], allow_blank=False)
        ws.add_data_validation(priority_val)
        priority_val.add(excel_columns_validation_ranges['1'])
    
    ws.title = title
    
    sheet_names = wb.sheetnames
    for s in sheet_names:
        ws = wb.get_sheet_by_name(s)
        if ws.dimensions == 'A1:A1':
            wb.remove_sheet(wb[s])

def save_excel():
    global zoom_scale, mapping_file, project_old, project_new
    try:
        if mapping_file == '':
            mapping_file = "Mappings '{}'-'{}'.xlsx".format(project_old, project_new)
        
        if os.path.exists(mapping_file) is True:
            overwrite_popup()
        
        for ws in wb.worksheets:
            ws.sheet_view.zoomScale = zoom_scale
            ws.auto_filter.ref = ws.dimensions
        wb.save(mapping_file)
        print("[END] Mapping file '{}' successfully generated.".format(mapping_file))
        print()
        sleep(2)
        exit()
    except Exception as e:
        print()
        print("[ERROR] ", e)
        os.system("pause")
        exit()


def get_minfields_issuetype(issue_details, all=0):
    min = 999
    i_types = {}
    for issuetype, fields in issue_details.items():
        mandatory_fields = set([field['id'] if field['required'] is True and field['id'] not in ['project', 'issuetype', 'summary'] else '' for field in fields.values()])
        if all == 0 and len(mandatory_fields) < min:
            min = len(mandatory_fields)
            min_type = issuetype
            min_fields = mandatory_fields
        else:
            mandatory_fields.remove('')
            i_types[issuetype] = mandatory_fields
    if all == 0:
        min_fields.remove('')
        return (min_type, min_fields)
    else:
        return i_types


def delete_extra_issues(max_processing_key):
    global start_jira_key, jira_old, jira_new, project_new, project_old, verbose_logging, delete_dummy_flag, threads
    
    def delete_issue(key):
        global username, password
        atlassian_jira_new = jira.Jira(JIRA_BASE_URL_NEW, username=username, password=password)
        atlassian_jira_new.delete_issue(key)
    
    # Calculating total Number of Issues in OLD JIRA Project
    jql_total_old = "project = '{}' AND key >= {} AND key <= {}".format(project_old, start_jira_key, max_processing_key)
    total_old = jira_old.search_issues(jql_total_old, startAt=0, maxResults=0, json_result=True)['total']
    
    # Calculating total Number of Migrated Issues to NEW JIRA Project
    jql_total_new = "project = '{}' AND summary !~ 'Dummy issue - for migration' AND key >= {} AND key <= {}".format(project_new, start_jira_key.replace(project_old, project_new), max_processing_key.replace(project_old, project_new))
    total_new = jira_new.search_issues(jql_total_new, startAt=0, maxResults=0, json_result=True)['total']
    
    print("[INFO] Total issues in Source Project: '{}' and total migrated issues: '{}'.".format(total_old, total_new))
    
    jql_total_new_for_deletion = "project = '{}' AND summary ~ 'Dummy issue - for migration' AND key >= {} AND key <= {}".format(project_new, start_jira_key.replace(project_old, project_new), max_processing_key.replace(project_old, project_new))
    total_new_for_deletion = jira_new.search_issues(jql_total_new_for_deletion, startAt=0, maxResults=0, json_result=True)['total']
    
    if delete_dummy_flag == 0:
        if total_old == total_new:
            if verbose_logging == 1:
                print("[INFO] Total 'dummy' issues to be deleted in new project: '{}'.".format(total_new_for_deletion))
            print("[START] 'Dummy' issue deletion is started. Please wait...")
            issues_for_delete = get_issues_by_jql(jira_new, jql_total_new_for_deletion, max_result=0)
            if issues_for_delete is not None:
                with concurrent.futures.ThreadPoolExecutor(max_workers=threads) as delete_executor:
                    future = [delete_executor.submit(delete_issue, i) for i in issues_for_delete]
            print("[END] 'Dummy' issues has been successfuly removed from target '{}' JIRA Project.".format(project_new), '', sep='\n')
        else:
            print("[ERROR] Not ALL issues have been migrated. 'Dummy' issues will not be removed to avoid any mapping issues.")
    else:
        print("[INFO] 'Dummy' issues will not be deleted due to 'Skip dummy deletion' flag was set.")


def create_dummy_issue(jira, project, issuetype, fields, old_issue):
    global issue_details_new

    new_data = {}
    new_data['project'] = project
    new_data['issuetype'] = eval('{"name": "' + issuetype + '"}')
    new_data['summary'] = "Dummy issue - for migration"

    for field in fields:
        for f in issue_details_new[issuetype]:
            if issue_details_new[issuetype][f]['id'] == field:
                default_value = issue_details_new[issuetype][f]['default value']
                allowed = issue_details_new[issuetype][f]['allowed values']
                type = issue_details_new[issuetype][f]['type']
                custom_type = issue_details_new[issuetype][f]['custom type']
                if type == 'option':
                    value = allowed[0] if default_value is None else default_value
                    new_data[field] = eval('{"value": "' + value + '"}')
                elif field in ['components', 'versions', 'fixVersions'] or custom_type == 'multiversion':
                    value = allowed[0] if default_value is None else default_value
                    new_data[field] = eval('[{"name": "' + value + '"}]')
                elif type == 'option-with-child':
                    new_data[field] = eval('{"value": "' + allowed[0][0] + '", "child": {"value": "' + allowed[0][1] + '"}}')
                elif type == 'string':
                    new_data[field] = 'Dummy' if default_value is None else default_value
                elif type == 'number':
                    new_data[field] = 0 if default_value is None else default_value
                elif type == 'array':
                    new_data[field] = ['Dummy'] if default_value is None else [default_value]
                else:
                    new_data[field] = default_value
    try:
        issue = jira.create_issue(fields=new_data)
    except Exception as e:
        print("[ERROR] Dummy issue cannot be created due to '{}'".format(e))
        return
    if int(old_issue.key.split('-')[1]) <= int(issue.key.split('-')[1]):
        return issue
    else:
        create_dummy_issue(jira, project, issuetype, fields, old_issue)


def convert_to_subtask(parent, new_issue, sub_task_id):
    """ This function will convert issue to sub-task via parsing HTML page and apply emulation of conversion via UI. """
    global auth
    
    session = requests.Session()
    
    url0 = JIRA_BASE_URL_NEW + '/secure/ConvertIssueSetIssueType.jspa?id=' + new_issue.id
    r = session.get(url=url0, auth=auth)
    soup = BeautifulSoup(r.text, features="lxml")
    try:
        guid = soup.find_all("input", type="hidden", id="guid")[0]['value']
    except:
        print("[ERROR] Issue can't be converted to Sub-Task")
        return
    
    url_11 = JIRA_BASE_URL_NEW + '/secure/ConvertIssueSetIssueType.jspa'
    payload_11 = {
        "parentIssueKey": parent,
        "issuetype": sub_task_id,
        "id": new_issue.id,
        "guid": guid,
        "Next >>": "Next >>",
        "atl_token": session.cookies.get('atlassian.xsrf.token'),
    }
    
    r = session.post(url=url_11, data=payload_11, headers={"Referer": url0})
    r.raise_for_status()
    
    url_12 = JIRA_BASE_URL_NEW + '/secure/ConvertIssueUpdateFields.jspa'
    payload_12 = {
        "id": new_issue.id,
        "guid": guid,
        "Next >>": "Next >>",
        "atl_token": session.cookies.get('atlassian.xsrf.token'),
    }
    
    r = session.post(url=url_12, data=payload_12)
    r.raise_for_status()
    
    url_13 = JIRA_BASE_URL_NEW + '/secure/ConvertIssueConvert.jspa'
    payload_13 = {
        "id": new_issue.id,
        "guid": guid,
        "Finish": "Finish",
        "atl_token": session.cookies.get('atlassian.xsrf.token'),
    }
    
    r = session.post(url=url_13, data=payload_13)
    r.raise_for_status()


def load_config(message=True):
    global mapping_file, JIRA_BASE_URL_OLD, JIRA_BASE_URL_NEW, project_old, project_new, last_updated_date, start_jira_key
    global team_project_prefix, old_board_id, default_board_name, temp_dir_name, limit_migration_data, threads
    
    if os.path.exists(config_file) is True:
        try:
            with open(config_file) as json_data_file:
                data = json.load(json_data_file)
            for k, v in data.items():
                if k == 'mapping_file':
                    mapping_file = v
                elif k == 'JIRA_BASE_URL_OLD':
                    JIRA_BASE_URL_OLD = v
                elif k == 'JIRA_BASE_URL_NEW':
                    JIRA_BASE_URL_NEW = v
                elif k == 'project_old':
                    project_old = v
                elif k == 'project_new':
                    project_new = v
                elif k == 'team_project_prefix':
                    team_project_prefix = v
                elif k == 'old_board_id':
                    old_board_id = v
                elif k == 'default_board_name':
                    default_board_name = v
                elif k == 'temp_dir_name':
                    temp_dir_name = v
                elif k == 'limit_migration_data':
                    limit_migration_data = v
                elif k == 'start_jira_key':
                    start_jira_key = v
                elif k == 'last_updated_date':
                    last_updated_date = v
                elif k == 'threads':
                    threads = v
            if message is True:
                print("[INFO] Configuration has been successfully loaded from '{}' file.".format(config_file))
        except Exception as er:
            print("[ERROR] Configuration file is corrupted. Default '{}' would be created instead.".format(config_file))
            print()
            save_config()
    else:
        print("[INFO] Config File not found. Default '{}' would be created.".format(config_file))
        print("[INFO] Migration configuration default values will be load from that file.")
        print()
        save_config()


def save_config(message=True):
    
    data = {'mapping_file': mapping_file,
            'JIRA_BASE_URL_OLD': JIRA_BASE_URL_OLD,
            'JIRA_BASE_URL_NEW': JIRA_BASE_URL_NEW,
            'project_old': project_old,
            'project_new': project_new,
            'team_project_prefix': team_project_prefix,
            'old_board_id': old_board_id,
            'default_board_name': default_board_name,
            'new_transitions': new_transitions,
            'temp_dir_name': temp_dir_name,
            'limit_migration_data': limit_migration_data,
            'start_jira_key': start_jira_key,
            'last_updated_date': last_updated_date,
            'threads': threads,
            }
    
    try:
        with open(config_file, 'w') as outfile:
            json.dump(data, outfile)
    except PermissionError as er:
        print("[ERROR] File '{}' has been opened for editing and can't be saved. Exception: {}".format(config_file, er))
        return
    
    if message is True:
        print("[INFO] Config file '{}' has been created.".format(config_file))
    else:
        print("[INFO] Config file '{}' has been updated.".format(config_file))
    print()


def update_new_issue_type(old_issue, new_issue, issuetype):
    global issue_details_old, issuetypes_mappings, sub_tasks, issue_details_new, create_remote_link_for_old_issue
    old_issuetype = old_issue.fields.issuetype.name
    
    def get_new_value_from_mapping(old_value, field_name):
        global field_value_mappings
        try:
            for new_value, old_values in field_value_mappings[field_name].items():
                if str(old_value) in old_values:
                    return new_value
        except:
            return old_value
    
    def get_old_system_field(new_field, old_issue=old_issue, old_issuetype=old_issuetype, new_issuetype=issuetype):
        global issue_details_old
        
        if new_field == 'Sprint':
            if issuetype in sub_tasks.keys():
                return None
            else:
                sprint_field = get_old_field('Sprint')
                issue_sprints = None if sprint_field is None else sprint_field
                if issue_sprints is not None:
                    new_issue_sprints = []
                    for sprint in issue_sprints:
                        name = ''
                        for attr in sprint[sprint.find('[')+1:-1].split(','):
                            if 'name=' in attr:
                                name = attr.split('name=')[1]
                            if name in new_sprints.keys():
                                new_issue_sprints.append(new_sprints[name]['id'])
                    if len(new_issue_sprints) == 0:
                        new_issue_sprints = None
                else:
                    new_issue_sprints = None
            return new_issue_sprints
        try:
            value = eval('old_issue.fields.' + issue_details_old[old_issuetype][new_field.strip()]['id'])
        except:
            return None
        if type(value) == list:
            cont_value = []
            for v in value:
                if hasattr(v, 'name'):
                    if ((issue_details_old[old_issuetype][new_field]['type'] == 'user' and jira_new.search_users(v) != [])
                            or issue_details_old[old_issuetype][new_field]['type'] != 'user'):
                        cont_value.append({"name": get_new_value_from_mapping(v.name, new_field)})
                elif hasattr(v, 'value'):
                    cont_value.append({"value": get_new_value_from_mapping(v.value, new_field)})
                else:
                    cont_value.append(get_new_value_from_mapping(v, new_field))
            return cont_value
        else:
            if hasattr(value, 'name'):
                if issue_details_old[old_issuetype][new_field]['type'] == 'user' and jira_new.search_users(value) == []:
                    return None
                else:
                    if new_field == 'Priority' and get_new_value_from_mapping(value.name, new_field) == '':
                        return {"name": issue_details_new[issuetype]['Priority']['default value']}
                    else:
                        return {"name": get_new_value_from_mapping(value.name, new_field)}
            elif hasattr(value, 'value'):
                return {"value": get_new_value_from_mapping(value.value, new_field)}
            else:
                if value is None and issue_details_old[old_issuetype][new_field]['type'] == 'string':
                    value = ''
                elif new_field in ['Epic Link', 'Parent Link']:
                    if value is None:
                        return None
                    else:
                        value = value.replace(project_old, project_new)
                        value = get_new_value_from_mapping(value, new_field)
                return value
    
    def get_old_field(new_field, old_issue=old_issue, old_issuetype=old_issuetype, new_issuetype=issuetype, data_val={}):
        global fields_mappings, issue_details_old, issue_details_new
        value = None
        concatenated_value = None
        
        def get_value(field, new_field=new_field, old_issue=old_issue, old_issuetype=old_issuetype):
            global issue_details_old, issue_details_new
            try:
                value = eval('old_issue.fields.' + issue_details_old[old_issuetype][field.strip()]['id'])
            except:
                try:
                    value = eval('old_issue.fields.' + field.strip())
                except:
                    value = None
            if issue_details_old[old_issuetype][field]['type'] == 'string' and issue_details_old[old_issuetype][field]['custom type'] == 'textfield':
                try:
                    value = value.replace('\n', '').replace('\t', ' ')
                except:
                    pass
            elif issue_details_old[old_issuetype][field]['type'] == 'number':
                try:
                    value = int(float(str(value).replace('\n', '').replace('\t', ' ')))
                except:
                    pass
            
            if issue_details_old[old_issuetype][field]['type'] in ['string', 'number', 'array'] and issue_details_new[new_issuetype][new_field]['custom type'] != 'com.atlassian.teams:rm-teams-custom-field-team':
                old_value = value
                if issue_details_new[new_issuetype][new_field]['type'] == 'option-with-child':
                    value = get_new_value_from_mapping(value, new_field)
                    try:
                        value_value = value.split(' --> ')[0]
                        value_child = value.split(' --> ')[1]
                        old_value = {"value": value_value, "child": {"value": value_child}}
                    except:
                        pass
            elif issue_details_old[old_issuetype][field]['type'] in ['option', 'user']:
                try:
                    old_value = value.value
                except:
                    try:
                        old_value = value.name
                    except:
                        old_value = value
            elif issue_details_old[old_issuetype][field]['type'] == 'option-with-child':
                if value is not None:
                    value_value = value.value
                    value_child = value.child.value
                    old_value = {"value": value_value, "child": {"value": value_child}}
            elif issue_details_new[new_issuetype][new_field]['custom type'] == 'com.atlassian.teams:rm-teams-custom-field-team':
                if issuetype in sub_tasks.keys():
                    return None
                else:
                    team = '' if value is None else get_team_id(value[0])
                    return team
            else:
                return get_new_value_from_mapping(value, new_field)
            
            return get_new_value_from_mapping(old_value, new_field)
        
        old_field = ''
        if new_field in fields_mappings[old_issuetype].keys():
            old_field = fields_mappings[old_issuetype][new_field]
        if old_field == '':
            try:
                old_field = [issue_details_old[old_issuetype][new_field.strip()].key()]
            except:
                return value
        if len(old_field) > 1:
            for o_field in old_field:
                if issue_details_new[new_issuetype][new_field]['type'] == 'string':
                    if concatenated_value is None:
                        if new_field != 'Description':
                            concatenated_value = ''
                        else:
                            try:
                                concatenated_value = data_val['description']
                            except:
                                concatenated_value = ''
                    added_value = '' if get_value(o_field) is None else get_str_from_lst(get_value(o_field))
                    if new_field == 'Description':
                        concatenated_value += '' if added_value == '' else '\\\\\\ [' + o_field + ']: ' + added_value
                    else:
                        concatenated_value += '' if get_str_from_lst(added_value) == '' else '[' + o_field + ']: ' + get_str_from_lst(added_value) + ' '
                elif issue_details_new[new_issuetype][new_field]['type'] == 'number':
                    if concatenated_value is None:
                        concatenated_value = 0
                    concatenated_value += 0 if get_value(o_field) is None else get_value(o_field)
                elif issue_details_new[new_issuetype][new_field]['type'] == 'array':
                    if concatenated_value is None:
                        if new_field != 'Labels':
                            concatenated_value = []
                        else:
                            concatenated_value = data_val['labels']
                    if issue_details_new[new_issuetype][new_field]['custom type'] == 'labels' or new_field == 'Labels':
                        concatenated_value.append('' if get_value(o_field) is None else get_value(o_field).replace(' ', '_').replace('\n', '_').replace('\t', '_'))
                    else:
                        concatenated_value.append('' if get_value(o_field) is None else get_value(o_field))
            value = concatenated_value
        else:
            value = get_value(old_field[0])
        return value
    
    def update_issuetype(issuetype, old_issutype, old_issue=old_issue):
        global issue_details_new
        data = {}
        mandatory_fields = get_minfields_issuetype(issue_details_new, all=1)
        data['issuetype'] = {'name': issuetype}
        data['summary'] = old_issue.fields.summary
        for field in mandatory_fields[old_issutype]:
            for f in issue_details_new[old_issutype]:
                if issue_details_new[old_issutype][f]['id'] == field:
                    default_value = issue_details_new[old_issutype][f]['default value']
                    allowed = issue_details_new[old_issutype][f]['allowed values']
                    type = issue_details_new[old_issutype][f]['type']
                    custom_type = issue_details_new[old_issutype][f]['custom type']
                    if type == 'option':
                        value = allowed[0] if default_value is None else default_value
                        data[field] = eval('{"value": "' + value + '"}')
                    elif field in ['components', 'versions', 'fixVersions'] or custom_type == 'multiversion':
                        value = allowed[0] if default_value is None else default_value
                        data[field] = eval('[{"name": "' + value + '"}]')
                    elif type == 'option-with-child':
                        data[field] = eval('{"value": "' + allowed[0][0] + '", "child": {"value": "' + allowed[0][1] + '"}}')
                    elif type == 'string':
                        data[field] = 'Dummy' if default_value is None else default_value
                    elif type == 'number':
                        data[field] = 0 if default_value is None else default_value
                    elif type == 'array':
                        data[field] = ['Dummy'] if default_value is None else [default_value]
                    else:
                        data[field] = default_value
        try:
            new_issue.update(notify=False, fields=data)
        except:
            new_issue.update(notify=True, fields=data)
    
    data_val = {}
    new_issuetype = new_issue.fields.issuetype.name
    # Checking for Sub-Task and convert to Sub-Task if necessary
    if issuetype in sub_tasks.keys():
        parent_field = old_issue.fields.parent
        parent = None if parent_field is None else parent_field.key.replace(project_old, project_new)
        if parent is not None and new_issuetype != issuetype:
            convert_to_subtask(parent, new_issue, sub_tasks[issuetype])
    data_val['summary'] = old_issue.fields.summary
    data_val['issuetype'] = {'name': issuetype}
    if new_issuetype != issuetype:
        update_issuetype(issuetype, new_issuetype)
    
    # System fields
    for n_field, n_values in issue_details_new[issuetype].items():
        if issuetype in sub_tasks.keys() and n_field in ['Sprint', 'Parent Link', 'Team']:
            continue
        if (n_values['custom type'] is None and n_field not in ['Issue Type', 'Summary', 'Project', 'Linked Issues', 'Attachment', 'Parent']) or (n_field in jira_system_fields):
            data_val[n_values['id']] = get_old_system_field(n_field)
    
    # Custom fields
    for n_field in fields_mappings[old_issuetype].keys():
        if n_field == '':
            continue
        if n_field not in issue_details_new[issuetype].keys() or n_field in ['Issue Type', 'Summary', 'Project', 'Linked Issues', 'Attachment', 'Parent'] or n_field in jira_system_fields:
            continue
        data_value = None
        o_field_value = get_old_field(n_field, data_val=data_val)
        n_field_value = '' if (o_field_value is None or o_field_value == 'None') else o_field_value
        if issue_details_new[issuetype][n_field]['type'] in ['number', 'date']:
            data_value = None if n_field_value == '' else n_field_value
        elif issue_details_new[issuetype][n_field]['type'] in ['string']:
            data_value = '' if n_field_value == '' else get_str_from_lst(n_field_value)
        elif issue_details_new[issuetype][n_field]['type'] in ['user', 'array']:
            if issue_details_new[issuetype][n_field]['custom type'] == 'multiuserpicker':
                if type(n_field_value) == list and n_field_value != '':
                    data_value = []
                    for i in n_field_value:
                        try:
                            data_value.append({"name": i.name})
                        except:
                            data_value.append({"name": None})
                else:
                    data_value = None if n_field_value == '' else [{"name": i} if i != '' else {"name": None} for i in n_field_value]
            elif issue_details_new[issuetype][n_field]['custom type'] == 'labels' or n_field == 'Labels':
                if type(n_field_value) == list and n_field_value != '':
                    data_value = None if n_field_value == '' else [i for i in n_field_value]
                else:
                    data_value = None if n_field_value == '' else [n_field_value]
            else:
                data_value = None if n_field_value == '' else {"name":  n_field_value}
        elif issue_details_new[issuetype][n_field]['type'] in ['option'] and issue_details_new[issuetype][n_field]['validated'] is True:
            data_value = None if n_field_value == '' else {"value":  n_field_value}
        elif issue_details_new[issuetype][n_field]['type'] == 'option-with-child' and n_field_value != '':
            try:
                data_value = {"value": n_field_value.value, "child": {"value": n_field_value.child.value}}
            except:
                data_value = None
        else:
            data_value = n_field_value
        
        data_val[issue_details_new[issuetype][n_field]['id']] = data_value

    if verbose_logging == 1:
        print("[INFO] The currently processing: '{}'".format(old_issue.key))
        print("[INFO] The details for update: '{}'".format(data_val))
    try:
        new_issue.update(notify=False, fields=data_val)
    except:
        new_issue.update(notify=True, fields=data_val)
    

def generate_template():
    global jira_old, jira_new, auth, username, password, project_old, project_new, mapping_file, JIRA_BASE_URL_NEW
    global JIRA_BASE_URL_OLD, issue_details_old, issue_details_new, migrate_statuses_check, threads
    
    username = user.get()
    password = passwd.get()
    mapping_file = file.get()
    if mapping_file == '':
        mapping_file = 'Migration Template for {} project to {} project.xlsx'.format(project_old, project_new)
    else:
        mapping_file = mapping_file.split('.xls')[0] + '.xlsx'
    main.destroy()
    if project_old == '' or project_new == '' or JIRA_BASE_URL_NEW == '' or JIRA_BASE_URL_OLD == '':
        print("[ERROR] Please enter missing configuration parameters in the new window.")
        change_configs()
    if len(username) < 6 or len(password) < 3:
        print('[ERROR] JIRA credentials are required. Please enter them on new window.')
        jira_authorization_popup()
    else:
        auth = (username, password)
        if verbose_logging == 1:
            print("[INFO] A connection attempt to JIRA server is started.")
        try:
            jira_old = JIRA(JIRA_BASE_URL_OLD, auth=auth, logging=False, async_=True, async_workers=threads, max_retries=1)
            jira_new = JIRA(JIRA_BASE_URL_NEW, auth=auth, logging=False, async_=True, async_workers=threads, max_retries=1)
        except Exception as e:
            print("[ERROR] Login to JIRA failed. Check your Username and Password. Exception: '{}'".format(e))
            os.system("pause")
            exit()
    if project_old == '' or project_new == '' or JIRA_BASE_URL_NEW == '' or JIRA_BASE_URL_OLD == '':
        print("[ERROR] Configuration parameters are not set. Exiting...")
        os.system("pause")
        exit()
    print('[START] Template is being generated. Please wait...')
    print()
    print("[START] Fields configuration downloading from '{}' and '{}' projects".format(project_old, project_new))
    issue_details_old = get_fields_list_by_project(jira_old, project_old)
    issue_details_new = get_fields_list_by_project(jira_new, project_new)
    print("[END] Fields configuration successfully processed.", '', sep='\n')
    if migrate_statuses_check == 1:
        get_transitions(project_new, JIRA_BASE_URL_NEW, new=True)
        get_transitions(project_old, JIRA_BASE_URL_OLD)
    get_hierarchy_config()
    
    for k, v in prepare_template_data().items():
        create_excel_sheet(v, k)
    save_excel()


def main_program():
    global jira_old, jira_new, auth, username, password, project_old, project_new, mapping_file, JIRA_BASE_URL_NEW
    global JIRA_BASE_URL_OLD, atlassian_jira_old, issue_details_old, issue_details_new, start_jira_key
    global limit_migration_data, verbose_logging, issuetypes_mappings, temp_dir_name, migrate_components_check
    global migrate_fixversions_check, validation_error, skip_migrated_flag, last_updated_date, updated_issues_num
    global create_remote_link_for_old_issue, threads, default_board_name, max_processing_key
    
    username = user.get()
    password = passwd.get()
    mapping_file = file.get().split('.xls')[0] + '.xlsx'
    
    # Checking the all mandatory fields are populated on Config page
    if validation_error == 1:
        change_configs()
    
    # Checking the Mapping File available
    if os.path.exists(mapping_file) is False or mapping_file == '.xlsx':
        load_file()
    main.destroy()

    if os.path.exists(mapping_file) is False or mapping_file == '.xlsx':
        print("[ERROR] Mapping File not found. Migration failed.")
        os.system("pause")
        exit()

    # Checking the JIRA credentials
    if len(username) < 6 or len(password) < 3:
        print('[ERROR] JIRA credentials are required. Please enter them on new window.')
        jira_authorization_popup()
    else:
        auth = (username, password)
        try:
            jira_old = JIRA(JIRA_BASE_URL_OLD, auth=auth, logging=False, async_=True, async_workers=threads, max_retries=1)
            jira_new = JIRA(JIRA_BASE_URL_NEW, auth=auth, logging=False, async_=True, async_workers=threads, max_retries=1)
            if create_remote_link_for_old_issue == 1:
                atlassian_jira_old = jira.Jira(JIRA_BASE_URL_OLD, username=username, password=password)
        except Exception as e:
            print("[ERROR] Login to JIRA failed. JIRA is unavailable or credentials are invalid. Exception: '{}'".format(e))
            os.system("pause")
            exit()
    
    # Starting Program
    print('[START] Migration process has been started. Please wait...')
    print()
    read_excel(file_path=mapping_file)
    print("[START] Fields configuration downloading from '{}' and '{}' projects".format(project_old, project_new))
    issue_details_old = get_fields_list_by_project(jira_old, project_old)
    issue_details_new = get_fields_list_by_project(jira_new, project_new)
    print("[END] Fields configuration successfully processed.", '', sep='\n')
    
    if migrate_statuses_check == 1:
        get_transitions(project_new, JIRA_BASE_URL_NEW, new=True)
    
    get_hierarchy_config()
    
    # Calculating the highest level of available Key in OLD project
    start_jira_key = project_old + '-' + str(start_jira_key)
    max_processing_key = project_old + '-' + str(int(limit_migration_data) + int(start_jira_key.split('-')[1]))

    # Check already migrated issues
    if skip_migrated_flag == 1:
        jql_last_migrated = "project = '{}' AND summary !~ 'Dummy issue - for migration' order by key desc".format(project_new)
        issues = get_issues_by_jql(jira_new, jql_last_migrated, max_result=1)
        last_migrated_id = issues[0] if issues is not None else None
        if last_migrated_id is None:
            last_migrated_id = project_new + '-1'
        new_start_jira_key = project_old + '-' + str(last_migrated_id.split('-')[1])
        if start_jira_key < new_start_jira_key:
            start_jira_key = new_start_jira_key
        if start_jira_key > max_processing_key:
            start_jira_key = max_processing_key
        print("[INFO] The first processing issue would be '{}'.".format(start_jira_key))
    
    # Add last updated issues to migration / update process
    if last_updated_date not in ['YYYY-MM-DD', '']:
        try:
            jql_latest = "project = '{}' AND key <= '{}' AND updated >= {}".format(project_old, max_processing_key, last_updated_date)
            updated_issues = get_issues_by_jql(jira_old, jql_latest, max_result=0)
            if updated_issues is not None:
                for i in updated_issues:
                    issue = jira_old.issue(i)
                    if issue.fields.issuetype.name not in items_lst.items():
                        items_lst[issue.fields.issuetype.name] = set()
                    items_lst[issue.fields.issuetype.name].add(issue.key)
        except:
            print("[ERROR] The value for Last Updated '{}' not in correct 'YYYY-MM-DD' format.".format(last_updated_date))
    
    if limit_migration_data != 0:
        jql_max = "project = '{}' AND key <= '{}' AND key >= '{}' order by key desc".format(project_old, max_processing_key, start_jira_key)
    else:
        jql_max = 'project = {} order by key desc'.format(project_old)
        max_processing_key = project_old + '-' + str(int(jira_old.search_issues(jql_max, startAt=0, maxResults=0, json_result=True)['total']))
    
    try:
        max_id = get_issues_by_jql(jira_old, jql_max, max_result=1)[0]
    except:
        max_id = max_processing_key
    if verbose_logging == 1:
        print("[INFO] The Maximum JIRA Key for OLD '{}' project to be migrated is '{}'.".format(project_old, max_id))
    
    # Components Migration
    if migrate_components_check == 1:
        migrate_components()
    
    # FixVersions Migration
    if migrate_fixversions_check == 1:
        migrate_versions()
    
    # Teams Migration (skipping if no mapping to Portfolio Teams)
    if migrate_teams_check == 1:
        for f_mappings in fields_mappings.values():
            if 'Team' in f_mappings.keys():
                get_all_shared_teams()
                break
    else:
        for issuestype, fields in fields_mappings.items():
            if 'Team' in fields.keys():
                fields_mappings[issuestype].pop('Team', None)
    
    # Creating / Cleaning Folder for Attachments migration
    if migrate_attachments_check == 1:
        create_temp_folder(temp_dir_name)
    
    # Sprints migration check
    if migrate_sprints_check == 1:
        if old_board_id == 0:
            migrate_sprints(proj_old=project_old, project=project_new, name=default_board_name)
        else:
            migrate_sprints(proj_old=project_old, board_id=old_board_id, project=project_new, name=default_board_name)
    else:
        if limit_migration_data != 0:
            jql_details = 'project = {} AND key >= {} AND key <= {} order by key ASC'.format(project_old, start_jira_key, max_processing_key)
        else:
            jql_details = 'project = {} AND key >= {} order by key ASC'.format(project_old, start_jira_key)
        get_issues_by_jql(jira_old, jql=jql_details, details=True, issue_details=issue_details_old)
    
    # Extra Logging
    if verbose_logging == 1:
        number_of_migrated = 0
        for v in items_lst.values():
            number_of_migrated += len(v)
        print("[INFO] The Number of issues to be migrated: {}".format(number_of_migrated))
        print('[INFO] The list of migrated issues by type:', items_lst)
    
    # Main Migration block
    for i in range(4):
        for k, v in issuetypes_mappings.items():
            if v['hierarchy'] == str(i):
                if k in items_lst.keys():
                    print("[INFO] The total number of '{}' issuetype: {}".format(k, len(items_lst[k])))
                migrate_issues(issuetype=k)

    # Cleaning Folder for Attachments migration
    if migrate_attachments_check == 1:
        clean_temp_folder(temp_dir_name)
    
    # Update and Close Sprints - after migration of issues are done
    if migrate_sprints_check == 1:
        migrate_sprints(proj_old=project_old, param='ACTIVE')

        # Calculating total Number of Issues in OLD JIRA Project
        jql_total_old = "project = '{}'".format(project_old)
        total_old = jira_old.search_issues(jql_total_old, startAt=0, maxResults=0, json_result=True)['total']
    
        # Calculating total Number of Migrated Issues to NEW JIRA Project
        jql_total_new = "project = '{}' AND summary !~ 'Dummy issue - for migration'".format(project_new)
        total_new = jira_new.search_issues(jql_total_new, startAt=0, maxResults=0, json_result=True)['total']
        
        if total_old == total_new:
            migrate_sprints(proj_old=project_old, param='CLOSED')
        else:
            print("[WARNING] Not ALL issues have been migrated from '{}' project. Remaining Issues: '{}'. Sprints will not be CLOSED until ALL issues migrated.".format(project_old, int(total_old) - int(total_new)))
            print()
        
    # Delete issues with Summary = 'Dummy Issue'
    delete_extra_issues(max_processing_key)


def overwrite_popup():
    """Function which shows Pop-Up window with question about overriding Excel file, if it already exists"""
    global mapping_file
    
    def create_new():
        global mapping_file
        popup.destroy()
        popup.quit()
        time_format = "%Y-%m-%dT%H:%M:%S"
        now = datetime.datetime.strptime(str(datetime.datetime.utcnow().isoformat()).split(".", 1)[0], time_format)
        mapping_file = mapping_file.split('.xls')[0] + '_' + str(now).replace(':', '-').replace(' ', '_') + '_UTC.xlsx'
    
    def override():
        popup.destroy()
        popup.quit()
    
    popup = tk.Tk()
    popup.title("Override File?")
    
    l1 = tk.Label(popup, text="File '{}' already exist.".format(mapping_file), foreground="black", font=("Helvetica", 10), pady=4, padx=8)
    l1.grid(row=0, column=0, columnspan=2)
    l2 = tk.Label(popup, text="Do you want to override existing file OR create a new one?", foreground="black", font=("Helvetica", 10), pady=4, padx=8)
    l2.grid(row=1, column=0, columnspan=2)
    
    b1 = tk.Button(popup, text="Override", font=("Helvetica", 9, "bold"), command=override, width=20, heigh=2)
    b1.grid(row=2, column=0, pady=10, padx=8)
    b2 = tk.Button(popup, text="Create New", font=("Helvetica", 9, "bold"), command=create_new, width=20, heigh=2)
    b2.grid(row=2, column=1, pady=10, padx=8)
    
    tk.mainloop()


def change_configs():
    """Function which shows Pop-Up window with question about JIRA credentials, if not entered"""
    global JIRA_BASE_URL_OLD, project_old, JIRA_BASE_URL_NEW, project_new, start_jira_key, limit_migration_data
    global default_board_name, old_board_id, team_project_prefix, last_updated_date, threads
    
    def config_save():
        global JIRA_BASE_URL_OLD, project_old, JIRA_BASE_URL_NEW, project_new, start_jira_key, limit_migration_data
        global default_board_name, old_board_id, team_project_prefix, validation_error, last_updated_date, threads
        
        validation_error = 0
        
        JIRA_BASE_URL_OLD = source_jira.get()
        JIRA_BASE_URL_NEW = target_jira.get()
        project_old = source_project.get()
        project_new = target_project.get()
        
        start_jira_key = first_issue.get()
        limit_migration_data = migrated_number.get()
        default_board_name = new_board.get()
        old_board_id = old_board.get()
        team_project_prefix = new_teams.get()
        threads = threads_num.get().strip()
        last_updated_date = last_updated.get().strip()
        
        if last_updated_date == 'YYYY-MM-DD':
            last_updated_date = ''
        
        config_popup.destroy()
        
        try:
            JIRA_BASE_URL_OLD = str(JIRA_BASE_URL_OLD).strip()
        except:
            if JIRA_BASE_URL_OLD == '':
                print("[ERROR] Source JIRA URL is empty.")
            else:
                print("[ERROR] Source JIRA URL is invalid.")
            validation_error = 1
        if JIRA_BASE_URL_OLD == '':
            print("[ERROR] Source JIRA URL is empty.")
            validation_error = 1
        
        try:
            JIRA_BASE_URL_NEW = str(JIRA_BASE_URL_NEW).strip()
        except:
            if JIRA_BASE_URL_NEW == '':
                print("[ERROR] Target JIRA URL is empty.")
            else:
                print("[ERROR] Target JIRA URL is invalid.")
            validation_error = 1
        if JIRA_BASE_URL_NEW == '':
            print("[ERROR] Target JIRA URL is empty.")
            validation_error = 1
        
        try:
            project_old = str(project_old).strip()
        except:
            if project_old == '':
                print("[ERROR] Source JIRA Project Key is empty.")
            else:
                print("[ERROR] Source JIRA Project Key is invalid.")
            validation_error = 1
        if project_old == '':
            print("[ERROR] Source JIRA Project Key is empty.")
        
        try:
            project_new = str(project_new).strip()
        except:
            if project_new == '':
                print("[ERROR] Target JIRA Project Key is empty.")
            else:
                print("[ERROR] Target JIRA Project Key is invalid.")
            validation_error = 1
        if project_new == '':
            print("[ERROR] Target JIRA Project Key is empty.")
            validation_error = 1

        if start_jira_key == '':
            start_jira_key = 1
        try:
            start_jira_key = int(start_jira_key.strip())
            if start_jira_key < 1:
                start_jira_key = 1
        except:
            try:
                start_jira_key = str(start_jira_key.strip()).split('-')[1]
            except:
                print("[ERROR] Start Issue Key is invalid.")
        
        if limit_migration_data in ['', 'ALL']:
            limit_migration_data = 0
        try:
            limit_migration_data = int(limit_migration_data.strip())
            if limit_migration_data < 0:
                limit_migration_data = 0
                print("[ERROR] Number of Total migrated issues can't be NEGATIVE. Defaulted to '0' - 'ALL'.")
        except:
            print("[ERROR] Number of Total migrated issues is invalid. Default value for ALL would be used.")
            limit_migration_data = 0
        
        try:
            default_board_name = str(default_board_name.strip())
            if default_board_name == '' and migrate_sprints_check == 1:
                default_board_name = 'Shared Sprints'
                print("[ERROR] Board Name for migrated Sprints can't be empty. Default Name '{}' will be used instead.".format(default_board_name))
        except:
            print("[ERROR] New Board name for migrated Sprints is invalid.")
            if migrate_sprints_check == 1:
                validation_error = 1
        
        try:
            if old_board_id != '':
                old_board_id = int(old_board_id.strip())
        except:
            if migrate_sprints_check == 1:
                print("[ERROR] Board ID for Sprints migration from Source JIRA in invalid. By default ALL Sprints will be migrated.")
            old_board_id = 0
        
        try:
            team_project_prefix = str(team_project_prefix)
        except:
            print("[ERROR] Prefix for Team names is invalid. Default '[{}] ' will be used.".format(project_old))
            team_project_prefix = '[' + project_old + '] '
        
        try:
            threads = int(threads.strip())
        except:
            print("[ERROR] Parallel Threads number is invalid. Default '1' will be used.")
            threads = 1
        
        if validation_error == 1:
            print("[WARNING] Mandatory Config data is invalid or empty. Please check the Config data again.")
        save_config()
        config_popup.quit()
    
    def config_popup_close():
        config_popup.destroy()
        config_popup.quit()
    
    def check_similar(field, value):
        """ This function required for fixing same valu duplication issue for second Tk window """
        global JIRA_BASE_URL_OLD, project_old, JIRA_BASE_URL_NEW, project_new, start_jira_key, limit_migration_data
        global default_board_name, old_board_id, team_project_prefix, validation_error, last_updated_date, threads
        
        fields = {"JIRA_BASE_URL_OLD": JIRA_BASE_URL_OLD,
                  "project_old": project_old,
                  "JIRA_BASE_URL_NEW": JIRA_BASE_URL_NEW,
                  "project_new": project_new,
                  "start_jira_key": start_jira_key,
                  "limit_migration_data": limit_migration_data,
                  "default_board_name": default_board_name,
                  "old_board_id": old_board_id,
                  "team_project_prefix": team_project_prefix,
                  "validation_error": validation_error,
                  "last_updated_date": last_updated_date,
                  "threads": threads,
                  }
        for f, v in fields.items():
            if str(value) == str(v) and field != f:
                return check_similar(field, ' ' + str(value))
        else:
            return value
    
    config_popup = tk.Tk()
    config_popup.title("JIRA Migration Tool - Configuration")

    JIRA_BASE_URL_OLD = check_similar("JIRA_BASE_URL_OLD", JIRA_BASE_URL_OLD)
    
    tk.Label(config_popup, text="Source JIRA URL:", foreground="black", font=("Helvetica", 10), pady=7, padx=5, wraplength=150).grid(row=0, column=0, rowspan=1)
    source_jira = tk.Entry(config_popup, width=45, textvariable=JIRA_BASE_URL_OLD)
    source_jira.insert(END, JIRA_BASE_URL_OLD)
    source_jira.grid(row=0, column=1, columnspan=2, padx=8)

    project_old = check_similar("project_old", project_old)
    
    tk.Label(config_popup, text="Source Project Key:", foreground="black", font=("Helvetica", 10), pady=7, padx=5, wraplength=150).grid(row=0, column=3, rowspan=1)
    source_project = tk.Entry(config_popup, width=16, textvariable=project_old)
    source_project.insert(END, project_old)
    source_project.grid(row=0, column=3, columnspan=2, padx=7, stick=E)

    JIRA_BASE_URL_NEW = check_similar("JIRA_BASE_URL_NEW", JIRA_BASE_URL_NEW)

    tk.Label(config_popup, text="Target JIRA URL:", foreground="black", font=("Helvetica", 10), pady=7, padx=5, wraplength=200).grid(row=1, column=0, rowspan=1)
    target_jira = tk.Entry(config_popup, width=45, textvariable=JIRA_BASE_URL_NEW)
    target_jira.insert(END, JIRA_BASE_URL_NEW)
    target_jira.grid(row=1, column=1, columnspan=2, padx=8)

    project_new = check_similar("project_new", project_new)
    
    tk.Label(config_popup, text="Target Project Key:", foreground="black", font=("Helvetica", 10), pady=7, padx=5, wraplength=150).grid(row=1, column=3, rowspan=1)
    target_project = tk.Entry(config_popup, width=16, textvariable=project_new)
    target_project.insert(END, project_new)
    target_project.grid(row=1, column=3, columnspan=2, padx=7, stick=E)
    
    tk.Label(config_popup, text="____________________________________________________________________________________________________________").grid(row=2, columnspan=5)
    
    tk.Label(config_popup, text="Detailed Configuration for migration. Defaults are '0' or empty for ALL Sprints / Issues:", foreground="black", font=("Helvetica", 11, "italic"), padx=10, wraplength=500).grid(row=3, column=0, columnspan=5)

    start_jira_key = check_similar("start_jira_key", start_jira_key)
    
    tk.Label(config_popup, text="Start migration from (Issue Key or Number):", foreground="black", font=("Helvetica", 10), pady=7, padx=5, wraplength=300).grid(row=4, column=0, columnspan=2)
    first_issue = tk.Entry(config_popup, width=20, textvariable=start_jira_key)
    first_issue.insert(END, start_jira_key)
    first_issue.grid(row=4, column=2, columnspan=1, padx=8)
    
    if limit_migration_data == 0:
        limit_migration_data = 'ALL'
    limit_migration_data = check_similar("limit_migration_data", limit_migration_data)
    
    tk.Label(config_popup, text="Number for migration:", foreground="black", font=("Helvetica", 10), pady=7, padx=5, wraplength=200).grid(row=4, column=3)
    migrated_number = tk.Entry(config_popup, width=10, textvariable=limit_migration_data)
    migrated_number.delete(0, END)
    migrated_number.insert(0, limit_migration_data)
    migrated_number.grid(row=4, column=4, columnspan=1, padx=8)

    default_board_name = check_similar("default_board_name", default_board_name)
    
    tk.Label(config_popup, text="New Board name for migrated Sprints:", foreground="black", font=("Helvetica", 10), pady=7, padx=5, wraplength=250).grid(row=5, column=0, columnspan=2)
    new_board = tk.Entry(config_popup, width=20, textvariable=default_board_name)
    new_board.insert(END, default_board_name)
    new_board.grid(row=5, column=2, columnspan=1, padx=8)

    if old_board_id == 0:
        old_board_id = ''
    old_board_id = check_similar("old_board_id", old_board_id)
    
    tk.Label(config_popup, text="Sprints from Board ID only:", foreground="black", font=("Helvetica", 10), pady=7, padx=5, wraplength=200).grid(row=5, column=3)
    old_board = tk.Entry(config_popup, width=10, textvariable=old_board_id)
    old_board.delete(0, END)
    old_board.insert(0, old_board_id)
    old_board.grid(row=5, column=4, columnspan=1, padx=8)
    
    tk.Label(config_popup, text="Prefix for Team names, if migrated:", foreground="black", font=("Helvetica", 10), pady=7, padx=5, wraplength=250).grid(row=6, column=0, columnspan=2)
    new_teams = tk.Entry(config_popup, width=20, textvariable=team_project_prefix)
    new_teams.insert(END, team_project_prefix)
    new_teams.grid(row=6, column=2, columnspan=1, padx=8)

    threads = check_similar("threads", threads)
    
    tk.Label(config_popup, text="Parallel Threads:", foreground="black", font=("Helvetica", 10), pady=7, padx=5, wraplength=200).grid(row=6, column=3)
    threads_num = tk.Entry(config_popup, width=10, textvariable=threads)
    threads_num.delete(0, END)
    threads_num.insert(0, threads)
    threads_num.grid(row=6, column=4, columnspan=1, padx=8)
    
    if last_updated_date == '':
        last_updated_date = 'YYYY-MM-DD'

    last_updated_date = check_similar("last_updated_date", last_updated_date)
    
    tk.Label(config_popup, text="Force update issues changed after that date, i.e. 'last updated >=  :", foreground="black", font=("Helvetica", 10), pady=7, padx=8, wraplength=500).grid(row=7, column=0, columnspan=4)
    last_updated = tk.Entry(config_popup, width=15, textvariable=last_updated_date)
    last_updated.insert(END, last_updated_date)
    last_updated.grid(row=7, column=3, columnspan=2, padx=70, stick=W)
    
    tk.Label(config_popup, text="____________________________________________________________________________________________________________").grid(row=8, columnspan=5)
    
    tk.Button(config_popup, text='Cancel', font=("Helvetica", 9, "bold"), command=config_popup_close, width=20, heigh=2).grid(row=9, column=0, pady=8, padx=20, sticky=W, columnspan=3)
    tk.Button(config_popup, text='Save', font=("Helvetica", 9, "bold"), command=config_save, width=20, heigh=2).grid(row=9, column=2, pady=8, padx=20, sticky=E, columnspan=3)
    
    tk.mainloop()


def jira_authorization_popup():
    """Function which shows Pop-Up window with question about JIRA credentials, if not entered"""
    global auth, username, password, jira_old, jira_new, atlassian_jira_old, JIRA_BASE_URL_OLD, JIRA_BASE_URL_NEW
    
    def jira_save():
        global auth, username, password, jira_old, jira_new, atlassian_jira_old, JIRA_BASE_URL_OLD, JIRA_BASE_URL_NEW
        global threads
        
        username = user.get()
        password = passwd.get()
        if len(username) < 3 or len(password) < 3:
            print("Invalid JIRA credentials were entered!")
            os.system("pause")
            exit()
        auth = (username, password)
        jira_popup.destroy()
        
        if verbose_logging == 1:
            print("[INFO] A connection attempt to JIRA server is started.")
        try:
            jira_old = JIRA(JIRA_BASE_URL_OLD, auth=auth, logging=False, async_=True, async_workers=threads, max_retries=1)
            jira_new = JIRA(JIRA_BASE_URL_NEW, auth=auth, logging=False, async_=True, async_workers=threads, max_retries=1)
            atlassian_jira_old = jira.Jira(JIRA_BASE_URL_OLD, username=username, password=password)
        except Exception as e:
            print("[ERROR] Login to JIRA failed. JIRA could be unavailable or User credentials are invalid. Exception: '{}'".format(e))
            os.system("pause")
            exit()
        jira_popup.quit()
    
    def jira_cancel():
        jira_popup.destroy()
        jira_popup.quit()
        print("[ERROR] Invalid JIRA credentials were entered! Program exits...")
        os.system("pause")
        exit()
    
    jira_popup = tk.Tk()
    jira_popup.title("[AUTHORIZATION] JIRA credentials required")
    
    tk.Label(jira_popup, text="To Migrate issues please enter your Username / Password for JIRA access.", foreground="black", font=("Helvetica", 9), padx=10, wraplength=210).grid(row=1, column=0, rowspan=2)
    tk.Label(jira_popup, text="Username").grid(row=1, column=1, pady=5)
    tk.Label(jira_popup, text="Password").grid(row=2, column=1, pady=5)
    
    user = tk.Entry(jira_popup)
    user.grid(row=1, column=2, pady=5)
    passwd = tk.Entry(jira_popup, width=20, show="*")
    passwd.grid(row=2, column=2, pady=5)
    
    tk.Button(jira_popup, text='Cancel', font=("Helvetica", 9, "bold"), command=jira_cancel, width=20, heigh=2).grid(row=4, column=0, pady=8, padx=20, sticky=W, columnspan=2)
    tk.Button(jira_popup, text='OK', font=("Helvetica", 9, "bold"), command=jira_save, width=20, heigh=2).grid(row=4, column=1, pady=8, padx=20, sticky=E, columnspan=2)
    
    tk.mainloop()


def change_migrate_fixversions(*args):
    global migrate_fixversions_check
    migrate_fixversions_check = process_fixversions.get()


def change_migrate_components(*args):
    global migrate_components_check
    migrate_components_check = process_components.get()


def change_migrate_sprints(*args):
    global migrate_sprints_check
    migrate_sprints_check = process_sprints.get()


def change_migrate_attachments(*args):
    global migrate_attachments_check
    migrate_attachments_check = process_attachments.get()


def change_migrate_metadata(*args):
    global migrate_metadata_check
    migrate_metadata_check = process_metadata.get()


def change_migrate_comments(*args):
    global migrate_comments_check
    migrate_comments_check = process_comments.get()


def change_migrate_links(*args):
    global migrate_links_check
    migrate_links_check = process_links.get()


def change_migrate_statuses(*args):
    global migrate_statuses_check
    migrate_statuses_check = process_statuses.get()


def change_migrate_teams(*args):
    global migrate_teams_check
    migrate_teams_check = process_teams.get()


def change_logging(*args):
    global verbose_logging
    verbose_logging = process_logging.get()


def change_linking(*args):
    global create_remote_link_for_old_issue
    create_remote_link_for_old_issue = process_old_linkage.get()


def change_dummy(*args):
    global delete_dummy_flag
    delete_dummy_flag = process_dummy_del.get()


def change_migrated(*args):
    global skip_migrated_flag
    skip_migrated_flag = process_non_migrated.get()


# ------------------ MAIN PROGRAM -----------------------------------
print("[INFO] Program has started. Please DO NOT CLOSE that window.")
load_config()
print("[INFO] Please IGNORE any WARNINGS - the connection issues are covered by Retry logic.")
print()

main = tk.Tk()
Title = main.title("JIRA Migration Tool" + " v_" + current_version)

tk.Label(main, text="Mapping Template Generation", foreground="black", font=("Helvetica", 11, "italic", "underline"), pady=10).grid(row=0, column=0, columnspan=2, rowspan=2, sticky=W, padx=80)
tk.Label(main, text="Step 1", foreground="black", font=("Helvetica", 12, "bold", "underline"), pady=10).grid(row=0, column=0, columnspan=3, rowspan=2, sticky=W, padx=15)
tk.Button(main, text='Generate Template', font=("Helvetica", 9, "bold"), command=generate_template, width=20, heigh=2).grid(row=0, column=3, pady=0, rowspan=2)

tk.Label(main, text="_____________________________________________________________________________________________________________________________").grid(row=2, columnspan=4)

tk.Label(main, text="Mapping Template:", foreground="black", font=("Helvetica", 10), pady=7, padx=5, wraplength=150).grid(row=3, column=0, rowspan=1, padx=80, sticky=W, columnspan=1)
file = tk.Entry(main, width=77, textvariable=mapping_file)
file.insert(END, mapping_file)
file.grid(row=3, column=0, columnspan=3, sticky=E, padx=0)
tk.Button(main, text='Browse', command=load_file, width=15).grid(row=3, column=3, pady=3, padx=8)

tk.Label(main, text="Migration Configuration", foreground="black", font=("Helvetica", 11, "italic", "underline"), pady=10).grid(row=4, column=0, columnspan=4, sticky=W, padx=80)

tk.Label(main, text="Step 2", foreground="black", font=("Helvetica", 12, "bold", "underline"), pady=10).grid(row=4, column=0, columnspan=3, rowspan=1, sticky=W, padx=15)

process_fixversions = IntVar(value=migrate_fixversions_check)
Checkbutton(main, text="Migrate all fixVersions / Releases from Source JIRA.", font=("Helvetica", 9, "italic"), variable=process_fixversions).grid(row=5, sticky=W, padx=70, column=0, columnspan=3, pady=0)
process_fixversions.trace('w', change_migrate_fixversions)

process_components = IntVar(value=migrate_components_check)
Checkbutton(main, text="Migrate all Components from Source JIRA.", font=("Helvetica", 9, "italic"), variable=process_components).grid(row=6, sticky=W, padx=70, column=0, columnspan=3, pady=0)
process_components.trace('w', change_migrate_components)

process_sprints = IntVar(value=migrate_sprints_check)
Checkbutton(main, text="Migrate Sprints (specified in Configs) from Source JIRA (Agile Add-on).", font=("Helvetica", 9, "italic"), variable=process_sprints).grid(row=7, sticky=W, padx=70, column=0, columnspan=3, pady=0)
process_sprints.trace('w', change_migrate_sprints)

process_teams = IntVar(value=migrate_teams_check)
Checkbutton(main, text="Migrate Teams from Source JIRA (Portfolio Add-on).", font=("Helvetica", 9, "italic"), variable=process_teams).grid(row=8, sticky=W, padx=70, column=0, columnspan=3, pady=0)
process_teams.trace('w', change_migrate_teams)

process_metadata = IntVar(value=migrate_metadata_check)
Checkbutton(main, text="Migrate Metadata (field values) for Issues.", font=("Helvetica", 9, "italic"), variable=process_metadata).grid(row=9, sticky=W, padx=70, column=0, columnspan=3, pady=0)
process_metadata.trace('w', change_migrate_metadata)

process_attachments = IntVar(value=migrate_attachments_check)
Checkbutton(main, text="Migrate all Attachments from Source JIRA issues.", font=("Helvetica", 9, "italic"), variable=process_attachments).grid(row=10, sticky=W, padx=70, column=0, columnspan=3, pady=0)
process_attachments.trace('w', change_migrate_attachments)

process_comments = IntVar(value=migrate_comments_check)
Checkbutton(main, text="Migrate all Comments from Source JIRA issues.", font=("Helvetica", 9, "italic"), variable=process_comments).grid(row=11, sticky=W, padx=70, column=0, columnspan=3, pady=0)
process_comments.trace('w', change_migrate_comments)

process_links = IntVar(value=migrate_links_check)
Checkbutton(main, text="Migrate all Links from Source JIRA issues.", font=("Helvetica", 9, "italic"), variable=process_links).grid(row=12, sticky=W, padx=70, column=0, columnspan=3, pady=0)
process_links.trace('w', change_migrate_links)

process_statuses = IntVar(value=migrate_statuses_check)
Checkbutton(main, text="Update all Statuses / Resolutions from Source JIRA issues (Project Admin access required).", font=("Helvetica", 9, "italic"), variable=process_statuses).grid(row=13, sticky=W, padx=70, column=0, columnspan=3, pady=0)
process_statuses.trace('w', change_migrate_statuses)

tk.Button(main, text='Change Configuration', font=("Helvetica", 9, "bold"), state='active', command=change_configs, width=20, heigh=2).grid(row=7, column=3, pady=4, rowspan=3)

tk.Label(main, text="_____________________________________________________________________________________________________________________________").grid(row=14, columnspan=4)

tk.Label(main, text="Migration Process", foreground="black", font=("Helvetica", 11, "italic", "underline"), pady=10).grid(row=15, column=0, columnspan=4, sticky=W, padx=80)

tk.Label(main, text="Step 3", foreground="black", font=("Helvetica", 12, "bold", "underline"), pady=10).grid(row=15, column=0, columnspan=3, rowspan=1, sticky=W, padx=15)


tk.Label(main, text="For migration process please enter your Username / Password for JIRA(s) access", foreground="black", font=("Helvetica", 10), padx=10, wraplength=260).grid(row=16, column=0, rowspan=2, columnspan=3, sticky=W, padx=80)
tk.Label(main, text="Username", foreground="black", font=("Helvetica", 10)).grid(row=16, column=1, pady=5, columnspan=3, sticky=W, padx=20)
tk.Label(main, text="Password", foreground="black", font=("Helvetica", 10)).grid(row=17, column=1, pady=5, columnspan=3, sticky=W, padx=20)
user = tk.Entry(main)
user.grid(row=16, column=1, pady=5, sticky=W, columnspan=3, padx=100)
passwd = tk.Entry(main, width=20, show="*")
passwd.grid(row=17, column=1, pady=5, sticky=W, columnspan=3, padx=100)

tk.Button(main, text='Start JIRA Migration', font=("Helvetica", 9, "bold"), state='active', command=main_program, width=20, heigh=2).grid(row=16, column=3, pady=4, padx=10, rowspan=2)

tk.Label(main, text="_____________________________________________________________________________________________________________________________").grid(row=18, columnspan=4)

tk.Label(main, text="Additional Configuration", foreground="black", font=("Helvetica", 10, "italic", "underline"), pady=10).grid(row=19, column=0, columnspan=4, sticky=W, padx=300)

process_logging = IntVar(value=verbose_logging)
Checkbutton(main, text="Switch Verbose Logging ON for migration process.", font=("Helvetica", 9, "italic"), variable=process_logging).grid(row=20, column=0, sticky=W, padx=20, columnspan=3, pady=0)
process_logging.trace('w', change_logging)

process_dummy_del = IntVar(value=delete_dummy_flag)
Checkbutton(main, text="Skip deletion of dummy issues (for testing purposes).", font=("Helvetica", 9, "italic"), variable=process_dummy_del).grid(row=21, column=0, sticky=W, padx=20, columnspan=3, pady=0)
process_dummy_del.trace('w', change_dummy)

process_old_linkage = IntVar(value=create_remote_link_for_old_issue)
Checkbutton(main, text="Add Remote Links to Source Issues.", font=("Helvetica", 9, "italic"), variable=process_old_linkage).grid(row=20, column=1, sticky=W, padx=70, columnspan=3, pady=0)
process_old_linkage.trace('w', change_linking)

process_non_migrated = IntVar(value=skip_migrated_flag)
Checkbutton(main, text="Skip already migrated issues.", font=("Helvetica", 9, "italic"), variable=process_non_migrated).grid(row=21, column=1, sticky=W, padx=70, columnspan=3, pady=0)
process_non_migrated.trace('w', change_migrated)

tk.Button(main, text='Quit', font=("Helvetica", 9, "bold"), command=main.quit, width=20, heigh=2).grid(row=22, column=0, pady=8, columnspan=4, rowspan=2)

tk.Label(main, text="Author: Dmitry Elsakov", foreground="grey", font=("Helvetica", 8, "italic"), pady=10).grid(row=23, column=1, sticky=E, padx=20, columnspan=3)

tk.mainloop()
