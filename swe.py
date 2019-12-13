from jira import JIRA
import urllib3
import getpass
import datetime
import win32com.client
import os
from datetime import timedelta
import time

pwd = getpass.getpass(prompt='JIRA password: ', stream='')
urllib3.disable_warnings()
jira = JIRA(basic_auth=('kprokopiuk', pwd), options={'server': 'http://cspjira:8080/jira', 'verify': False})
# jira = JIRA(basic_auth=('kprokopiuk', '****'), options={'server': 'https://partnerjira.g2-networks.com', 'verify': False})
print('\nLogin successful!')


def search():
    try:
        os.system('cls')
        print('Search for issues in Jira\n')
        num = input('Input issue ID number(SDSFC): ')
        issue = jira.issue('SDSFC-'+num)

        print('Project:', issue.fields.project)
        print('ID: '+issue.key)
        print('Type: '+issue.fields.issuetype.name)
        print('Status:', issue.fields.status)
        print('Priority:', issue.fields.priority)
        print('Resolution:', issue.fields.resolution)
        # print('Components:', issue.fields.components[0])
        # print('Labels:', issue.fields.labels)
        print('Type: '+issue.fields.issuetype.name)
        print('Assignee:', issue.fields.assignee)
        print('Reporter:', issue.fields.reporter)
        print('Summary:', issue.fields.summary)
        print('Description:', issue.fields.description)
        print('Attachments:', issue.fields.attachment)
        print('Created:', issue.fields.created)
        print('Updated:', issue.fields.updated)
        print('--- End of file ---')
    except:
        print('Issue number SDSFC-'+num+' does not exist!')
    input('\nPress any key to continue...')


def last_issue():
    os.system('cls')
    print('Shows last created issue in Jira\n')
    # shows last issue, full info about last issue and full info about last 10 issues
    last_issue = jira.search_issues('project=SDSFC')[0]
    print('Last issue:', last_issue)
    # last_issue_full = jira.search_issues('project=SWLORO')[:1]
    # print('Last issue full info:', last_issue_full)
    # all_issues = jira.search_issues('project=SWLORO')[:11]
    # print('Last 10 issues full info:', all_issues)
    print('Project:', last_issue.fields.project)
    print('ID: ' + last_issue.key)
    print('Type: ' + last_issue.fields.issuetype.name)
    print('Priority:', last_issue.fields.priority)
    # print('Components:', last_issue.fields.components[0])
    print('Labels:', last_issue.fields.labels)
    print('Type: ' + last_issue.fields.issuetype.name)
    print('Status:', last_issue.fields.status)
    # print('Resolution:', last_issue.fields.resolution)
    print('Assignee:', last_issue.fields.assignee)
    print('Reporter:', last_issue.fields.reporter)
    print('Summary:', last_issue.fields.summary)
    print('Description:\n', last_issue.fields.description)
    print('Created:', last_issue.fields.created)
    print('Updated:', last_issue.fields.updated)

    input('\nPress any key to continue...')


def menu():
    os.system('cls')
    print('Jira App for SWEDEN project SDSFC')
    print('\nMain Menu:\n')
    print('1. Search')
    print('2. Latest Issue')
    print('\nE. Exit')
    choice = input('\nSelect Option: ')
    if choice == '1':
        search()
        menu()
    elif choice == '2':
        last_issue()
        menu()
    elif choice.lower() == 'e':
        input('\nPress any key to close...')
        exit(0)
    elif choice.lower() == 'exit':
        input('\nPress any key to close...')
        exit(0)
    else:
        input('\nWrong input...')
        os.system('cls')
        menu()


menu()
