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
jira = JIRA(basic_auth=('kprokopiuk', pwd), options={'server': 'https://partnerjira.g2-networks.com', 'verify': False})
# jira = JIRA(basic_auth=('kprokopiuk', '****'), options={'server': 'https://partnerjira.g2-networks.com', 'verify': False})
print('\nLogin successful!')


def search():
    try:
        os.system('cls')
        print('Search for issues in Jira\n')
        swloro = input('Input issue ID number(SWLORO): ')
        issue = jira.issue('SWLORO-'+swloro)

        # print('Project:', issue.fields.project)
        print('ID: '+issue.key)
        print('Type: '+issue.fields.issuetype.name)
        print('Priority:', issue.fields.priority)
        print('Components:', issue.fields.components[0])
        print('Labels:', issue.fields.labels)
        print('Type: '+issue.fields.issuetype.name)
        print('Status:', issue.fields.status)
        # print('Resolution:', issue.fields.resolution)
        print('Assignee:', issue.fields.assignee)
        print('Reporter:', issue.fields.reporter)
        print('Summary:', issue.fields.summary)
        print('Description:', issue.fields.description)
        print('Attachments:', issue.fields.attachment)
        print('Created:', issue.fields.created)
        print('Updated:', issue.fields.updated)
        print('--- End of file ---')
    except:
        print('Issue number SWLORO-'+swloro+' does not exist!')
    input('\nPress any key to continue . . . ')

# <JIRA Project: key='SWLOROTST', name='Test Swiss LoRo (only for testing)', id='13984'>,
# <JIRA Project: key='SWLORO', name='Swiss LoRo', id='13985'>


def create():
    os.system('cls')
    print('Creating ticket for last unread email in Bastion folder\n')
    now = datetime.datetime.utcnow()
    now2 = datetime.datetime.now()
    day_ago = now2 - timedelta(days=1)
    now2 = now.timestamp()
    day_ago = day_ago.timestamp()
    now = now - timedelta(hours=-1)
    now = now.strftime("%Y-%m-%dT%H:%M:%S.125+0100")
    # print(now)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    pre_inbox = outlook.GetDefaultFolder(6)
    inbox = pre_inbox.Folders(2)

    messages = inbox.Items
    # message = messages.GetLast()

    global flag_unread
    flag_unread = 0
    try:
        for message in messages:
            if message.receivedtime.timestamp() > day_ago:
                # print('Checking for new emails.')
                if message.Unread == True:
                    print('You are about to create ticket for email:\n')
                    print(message.creationtime.strftime("%Y-%m-%d %H:%M:%S"), 'from', message.sender, ':',
                          message.subject)
                    proceed = input('Continue? [Y/N]: ')
                    if proceed.lower() == 'y':
                        sender = message.Sender
                        sub_line = message.subject
                        body_content = message.body
                        flag_unread = 1
                        engineer_end = sub_line.find('@')
                        engineer = sub_line[39:engineer_end]
                        subject = sub_line[4:]

                        begin = body_content.find('A validation is')
                        end = body_content.find('Please follow this link to answer:')
                        body_content_cut = body_content[begin:end]

                        issue_dict = {
                            'project': 'SWLORO',
                            'issuetype': 'Incident',
                            'summary': subject,
                            'description': body_content_cut,
                            'components': [{'name': '-Access Wallix'}],
                            'customfield_10083': {'value': 'None'},
                            'labels': [engineer],
                            'customfield_10012': now,
                        }
                        print('Creating ticket')

                        # print(message.CreationTime)

                        jira.create_issue(fields=issue_dict)
                        # print(issue_dict)
                        print('Ticket created')

                        last_issue = jira.search_issues('project=SWLORO')[0]
                        issue = jira.issue(last_issue)
                        jira.transitions(issue)
                        # print(jira.transitions(issue))
                        reporter = 'kprokopiuk'
                        reporter2 = {'name': reporter}
                        print('Starting to investigate. Assignee changed to', issue.fields.reporter)
                        jira.transition_issue(issue, transition='Start Investigate', assignee=reporter2)
                        print('Adding comment: "Approved."')
                        jira.add_comment(issue, 'Approved.')
                        print('Issue resolved')
                        jira.transition_issue(issue, transition='Resolve', assignee=reporter2)
                        print('Issue closed')
                        jira.transition_issue(issue, transition='Close Issue')
                        print('Completed successfully')

                        create_link = input('Do you want to link this ticket to an existing one? [Y/N]: ')
                        try:
                            if create_link.lower() == 'y':
                                parent_number = input('Input issue ID number(SWLORO-): ')
                                print('Found existing ticket under number SWLORO-'+parent_number)
                                new_issue = jira.search_issues('project=SWLORO')[0]
                                # print(new_issue)
                                parent_issue = ('SWLORO-' + parent_number)
                                # print(parent_issue)
                                print('Linking', new_issue, 'to', parent_issue)
                                jira.create_issue_link('relates to', new_issue, parent_issue, None)
                                print(new_issue, 'issue linked to', parent_issue)
                            if create_link.lower() == 'n':
                                pass
                        except:
                            message.Unread = False
                            print('Unspecified exception, please link tickets manually!')
                            input('Press any key to continue . . . ')
                            menu()
                        message.Unread = False
                        print('Returning to menu...')
                    else:
                        print('Skipped')
                        pass
        flag_unread = 2
        print('No more unread emails!')
    except:
        print('There are no unread emails in Bastion')
    if flag_unread == 0:
        print('There are no unread emails in Bastion')
    input('\nPress any key to continue . . . ')


def last_issue():
    os.system('cls')
    print('Shows last created issue in Jira\n')
    # shows last issue, full info about last issue and full info about last 10 issues
    last_issue = jira.search_issues('project=SWLORO')[0]
    print('Last issue:', last_issue)
    # last_issue_full = jira.search_issues('project=SWLORO')[:1]
    # print('Last issue full info:', last_issue_full)
    # all_issues = jira.search_issues('project=SWLORO')[:11]
    # print('Last 10 issues full info:', all_issues)
    # print('Project:', last_issue.fields.project)
    print('ID: ' + last_issue.key)
    print('Type: ' + last_issue.fields.issuetype.name)
    print('Priority:', last_issue.fields.priority)
    print('Components:', last_issue.fields.components[0])
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

    input('\nPress any key to continue . . . ')


def issues_list():
    last_issue = jira.search_issues('project=SWLORO')[0]
    print('Last issue:', last_issue)
    last_issue_full = jira.search_issues('project=SWLORO')[:1]
    print('Last issue full info:', last_issue_full)
    all_issues = jira.search_issues('project=SWLORO')[:11]
    print('Last 10 issues full info:', all_issues)
    menu()


def listen_for_issues():
    now = datetime.datetime.now()
    day_ago = now - timedelta(days=1)
    now = now.timestamp()
    day_ago = day_ago.timestamp()
    # print('now:     ', now)
    # print('day ago: ', day_ago)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    pre_inbox = outlook.GetDefaultFolder(6)
    inbox = pre_inbox.Folders(2)

    # print(inbox)

    messages = inbox.Items
    message = messages.GetLast()
    sender = message.Sender
    sub_line = message.subject
    body_content = message.body

    # datetime.datetime.now().strftime('%H:%M:%S')

    os.system('cls')
    print('Auto Jira App started')
    print('Waiting for new email (Press Ctrl+C to abort)')

    while True:
        try:
            for message in messages:
                if message.receivedtime.timestamp() > day_ago:
                    # print('Checking for new emails.')
                    if message.Unread == True:
                        print(datetime.datetime.now().strftime('%H:%M:%S'), ': Found new unread email in Bastion.')
                        print(message.creationtime.strftime("%Y-%m-%d %H:%M:%S"), 'from', message.sender, ':',
                              message.subject)
                        now = datetime.datetime.utcnow()
                        now = now - timedelta(hours=-1)
                        now = now.strftime("%Y-%m-%dT%H:%M:%S.125+0100")
                        # print(now)

                        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                        pre_inbox = outlook.GetDefaultFolder(6)
                        inbox = pre_inbox.Folders(2)

                        messages = inbox.Items
                        # message = messages.GetLast()
                        sender = message.Sender
                        sub_line = message.subject
                        body_content = message.body

                        engineer_end = sub_line.find('@')
                        engineer = sub_line[39:engineer_end]
                        subject = sub_line[4:]

                        begin = body_content.find('A validation is')
                        end = body_content.find('Please follow this link to answer:')
                        body_content_cut = body_content[begin:end]

                        issue_dict = {
                            'project': 'SWLORO',
                            'issuetype': 'Incident',
                            'summary': subject,
                            'description': body_content_cut,
                            'components': [{'name': '-Access Wallix'}],
                            'customfield_10083': {'value': 'None'},
                            'labels': [engineer],
                            'customfield_10012': now,
                        }
                        print('Creating ticket')

                        # print('Time:', message.CreationTime)

                        jira.create_issue(fields=issue_dict)

                        # print(issue_dict)
                        print('Ticket created')

                        last_issue = jira.search_issues('project=SWLORO')[0]
                        issue = jira.issue(last_issue)
                        jira.transitions(issue)
                        # print(jira.transitions(issue))
                        reporter = 'kprokopiuk'
                        reporter2 = {'name': reporter}
                        print('Starting to investigate. Assignee changed to', issue.fields.reporter)
                        jira.transition_issue(issue, transition='Start Investigate', assignee=reporter2)
                        print('Adding comment: "Approved."')
                        jira.add_comment(issue, 'Approved.')
                        print('Issue resolved')
                        jira.transition_issue(issue, transition='Resolve', assignee=reporter2)
                        print('Issue closed')
                        jira.transition_issue(issue, transition='Close Issue')
                        print('Completed successfully')
                        message.Unread = False
                    else:
                        # print("Didn't find any new emails. Going to sleep for 60 seconds.")
                        time.sleep(60)
        except KeyboardInterrupt:
            print('Auto Jira task aborted')
            input('\nPress any key to continue . . . ')
            menu()

        except:
            # print('No more unread emails. Going to sleep for 60 seconds.')
            time.sleep(60)


def one_time_check():
    os.system('cls')
    print('Looking for unread emails in Bastion folder in Outlook\n')
    now = datetime.datetime.now()
    day_ago = now - timedelta(days=1)
    now = now.timestamp()
    day_ago = day_ago.timestamp()
    # print('now:     ', now)
    # print('day ago: ', day_ago)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    pre_inbox = outlook.GetDefaultFolder(6)
    inbox = pre_inbox.Folders(2)

    # print(inbox)

    messages = inbox.Items
    message = messages.GetLast()
    sender = message.Sender
    sub_line = message.subject
    body_content = message.body

    i = 0
    try:
        for message in messages:
            if message.receivedtime.timestamp() > day_ago:
                # print('Checking for new emails.')
                if message.Unread == True:
                    # print(datetime.datetime.now().strftime('%H:%M:%S'), ': There is unread email in Bastion')
                    i += 1
                    print(message.creationtime.strftime("%Y-%m-%d %H:%M:%S"), 'from', message.sender, ':',
                          message.subject)
    except:
        pass
    print('')
    if i == 0:
        print('There are', i, 'unread emails in Bastion')
    elif i == 1:
        print('There is', i, 'unread email in Bastion')
    else:
        print('There are', i, 'unread emails in Bastion')
    input('\nPress any key to continue . . . ')


def create_single_link():
    try:
        os.system('cls')
        print('Creating single link between two tickets in Jira\n')
        new_issue = input('Input issue ID number you want to link(SWLORO-): ')
        parent_number = input('Input issue ID number you want it to be linked to(SWLORO-): ')
        print('Found existing ticket under number SWLORO-' + parent_number)
        new_issue = f'SWLORO-{new_issue}'
        parent_issue = ('SWLORO-' + parent_number)
        print('Linking', new_issue, 'to', parent_issue)
        jira.create_issue_link('relates to', new_issue, parent_issue, None)
        print(new_issue, 'issue linked to', parent_issue)
    except:
        print('Error. Please try again')
    input('\nPress any key to continue . . . ')


def menu():
    os.system('cls')
    print('Jira App for Wallix approvals (project SWLOROTST)')
    print('\nMain Menu:\n')
    print('1. Search')
    print('2. Latest Issue')
    print('3. Check email')
    print('4. Create single ticket')
    print('5. Link single ticket')
    print('6. Auto Jira App')
    print('\nE. Exit')
    choice = input('\nSelect Option: ')
    if choice == '1':
        search()
        menu()
    elif choice == '2':
        last_issue()
        menu()
    elif choice == '3':
        one_time_check()
        menu()
    elif choice == '4':
        create()
        menu()
    elif choice == '5':
        create_single_link()
        menu()
    elif choice == '6':
        listen_for_issues()
        menu()
    elif choice.lower() == 'e':
        input('\nPress any key to close . . . ')
        exit(0)
    elif choice.lower() == 'exit':
        input('\nPress any key to close . . . ')
        exit(0)
    else:
        input('\nWrong input . . . ')
        os.system('cls')
        menu()


menu()
