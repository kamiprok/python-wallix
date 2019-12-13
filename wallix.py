import email
import imaplib

username = 'email'
password = 'password'

mail = imaplib.IMAP4_SSL('outlook.office365.com')
mail.login(username, password)

result, mailboxes = mail.list()
mail.select('Inbox')

result, numbers = mail.uid('search', None, 'ALL')
uids = numbers[0].split()

# retrieve the headers (without setting the 'seen' flag) of the last 10 messages
# in the list of UIDs
result, messages = mail.uid('fetch', ','.join(uids[-10:]), '(BODY.PEEK[HEADER])')

# Convert the messages into email message object and print out the sender and
# the subject.
for _, message in messages[::2]:
    msg = email.message_from_string(message)
    print('{}: {}'.format(msg.get('from'), msg.get('subject')))

