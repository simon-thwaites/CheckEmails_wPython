# WORKING as get_info_from_inbox.py
# ======================================================================
# This script will scan opened Outlook inbox for web submissions and:
#     - extract information from the body of the email
#     - append these values to a csv file
#     - send email to the applicant
#     - move the email to the destination folder
#     - if Outlook isn't running, this will open Outlook before running rest of script.
# ======================================================================
# author: Simon Thwaites
# Email:  simonthwaites1991@gmail.com
# Date:   17/12/2019
# ======================================================================

import csv
import win32com.client
from win32com.client import Dispatch, constants
import time

# check to see if outlook is running, if not open
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

# initialise
writeRows = []
submissionCounter = 0

# loop through inbox in reverse order because can always access [first] when moving emails. 
# Otherwise might encounter error from out of index
for x in range(lengthInbox, 0, -1):
    message = messages[x]               # outlook indexing starts at 1 not 0
    subject_content = message.Subject   # email subject line
    sender = message.SenderEmailAddress # email sender
    
    # check email sender and email subject both fit criteria
    if subject_content == 'Webform submission - Knee function study' and sender == 'web.requests@adelaide.edu.au':
        submissionCounter += 1
        message.UnRead = True            # mark email as unread
        body_content = message.Body      # email body text
        lengthBody = len(body_content)
        
        # extract indexes of subject fields
        fullName_index = body_content.find('Full name')
        age_index = body_content.find('Age')
        gender_index = body_content.find('Gender')
        phone_index = body_content.find('Phone number')
        email_index = body_content.find('Email address')
        height_index = body_content.find('Height')
        weight_index = body_content.find('Weight')
        
        # extract text elements of the subjects fields
        subject_fullName = body_content[fullName_index+9:age_index-1]
        subject_age = body_content[age_index+3:gender_index-1]
        subject_gender = body_content[gender_index+6:phone_index-1]
        subject_phone = body_content[phone_index+12:email_index-1]
        subject_email = body_content[email_index+13:height_index-1]
        subject_height = body_content[height_index+6:weight_index-1]
        subject_weight = body_content[weight_index+6:-1]
        
        # remove leading and trailing whitespace
        subject_fullName = subject_fullName.strip()
        subject_age = subject_age.strip()
        subject_gender = subject_gender.strip()
        subject_phone = subject_phone.strip()   # note that excel will remove leading zeros
        subject_email = subject_email.strip()
        subject_height = subject_height.strip()
        subject_weight = subject_weight.strip()
        
        # also get the date and time the email was recieved
        date_index = body_content.find(",")
        time_index = body_content.find("-")
        date_received = body_content[date_index+2:time_index-1]
        time_received = body_content[time_index+2:time_index+7]
        
        # setup for sending emails
        const = win32com.client.constants
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = "RE: Webform submission - Knee function study"
        newMail.To = subject_email
        newMail.Body = """Hi """+subject_fullName+""",

Thank you for expressing interest in the knee function study! This research is investigating a particular type of leg fracture and we require some healthy control subjects. You have been added to a database of potential healthy volunteers for this study. 

Your inclusion into this study works on a best-matched scenario, that is: if someone sustains the fracture that we are investigating, we will look to recruit the best-matched healthy control available (for age, sex, height, and body mass). This means that even though you have expressed interest in the study, it may take some time for you to be matched with a fracture patient, or you may not be matched at all. 

If you are matched, you will be contacted to book in for your testing session in The Clinical Research Facility at the Adelaide Health and Medical Sciences building at a time that suits you. 

Thanks again for expressing your interest in participating in this study.

Kind regards,

Simon Thwaites
Ph.D. Candidate
Faculty of Health & Medical Sciences
Adelaide Medical School
Level 7 | AHMS Building | The University of Adelaide | SA | 5000 
Tel: (08) 8313 0088
Email: simon.thwaites@adelaide.edu.au
 
CRICOS Provider Number 00123M


-----------------------------------------------------------
IMPORTANT: This message may contain confidential or legally privileged information. If you think it was sent to you by mistake, please delete all copies and advise the sender. For the purposes of the SPAM Act 2003, this email is authorised by The University of Adelaide.

Think green: read on the screen.
"""
        # send the message
        newMail.Send()
        
        # generate the new row of data
        newRow = [subject_fullName,subject_age,subject_gender,subject_phone,subject_email,\
                  subject_height,subject_weight,date_received,time_received]
        
        # append all new rows to be written to the cv file
        writeRows.append(newRow)
            
        # move the message to destination folder    
        message.Move(destination)

# now append the new row of data to fullFile with 'a', remove blank lines with newline = ''
# but reverse this order again so that the oldest submission is appended first
for x in range(submissionCounter-1,-1,-1):
    with open(fullFile, 'a', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(writeRows[x])
        f.close()