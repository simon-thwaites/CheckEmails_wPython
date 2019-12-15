# now check to see if outlook is running, if not open
import time
import csv
import win32com.client
from win32com.client import Dispatch, constants

def outlook_is_running():
    import win32ui
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
        return True
    except win32ui.error:
        return False

if not outlook_is_running():
    import os
    os.startfile("outlook")

# now wait 30 secs for outlook to open
time.sleep(30)
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# directory of email datbase and file name
dir = 'C:/Users/a1194788/Box/01. PhD/05. Participant Information/02. Prospective healthy cohort/'
fileName = 'Healthy-Cohort_email-database.csv'
fullFile = dir + fileName

# define Inbox and destination folders in outlook
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
lengthInbox = len(messages)
destination = outlook.GetDefaultFolder(6).Folders['PhD'].Folders['Particpants'].Folders['UoA web enquiries'].Folders['First reply sent']

print(lengthInbox)

# initialise
writeRows = []
submissionCounter = 0

for x in range(lengthInbox, 0, -1):
    message = messages[x]
    subject_content = message.Subject   # email subject line
    print(subject_content)