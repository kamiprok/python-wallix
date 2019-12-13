import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
pre_inbox = outlook.GetDefaultFolder(6)
inbox = pre_inbox.Folders(2)

print(inbox)

messages = inbox.Items
message = messages.GetLast()
sender = message.Sender
sub_line = message.subject
body_content = message.body

engineer_end = sub_line.find('@')
engineer = sub_line[39:engineer_end]
subject = sub_line[4:]

begin = body_content.find('A validation is')
end = body_content.find('Please follow this link to answer:')

print(message.CreationTime)
print('sender:', sender)
print('Engineer:', engineer)
print('Subject:', subject)
print('Body:', body_content[begin:end])

#check if message.CreationTime is in list, if it is do nothing, if not make ticket and add it to the list