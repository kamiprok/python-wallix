import win32com.client
import re
import webbrowser
from selenium import webdriver

outlook = win32com.client.Dispatch("Outlook.Application")

'''inbox = outlook.Folder[1] # "6" refers to the index of a folder

messages = inbox.Items
message = messages.GetLast()
sender = message.Sender
sub_line = message.subject
body_content = message.body

print('sender: ', sender)
print('Subject: ', sub_line)
print('Body: ', body_content)
'''
for i in outlook.Folders:
    print(i)

inbox = outlook.Folders[2]
print(inbox)

'''
urls = re.findall(r'(https?://\S+)', body_content)
print(urls)

url = urls[0]
webbrowser.open(url)

browser	=	webdriver.Chrome('C:/Users/kprokopiuk/Downloads/chromedriver.exe')
browser.get(url)
userElem	=	browser.find_element_by_id('user_name')
userElem.send_keys('login') #admn no here
passwordElem	=	browser.find_element_by_id('passwd')
passwordElem.send_keys('pw') # password here
loginElem	=	browser.find_element_by_id('submitLogin')
loginElem.click()
'''
