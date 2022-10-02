# CheckEmails_wPython
(working on python 3.7.4)

This script is very specific, but feel free to adapt to your needs.

get_info_from_inbox.py will check your Outlook inbox for emails with a specific subject and sender email address.
This script will first check that Outlook is open, then extracts information from a standardised web submission form (see below), appends this information to a .csv file, generates an automatic reply to the submitted email address, then moves the email to a destination folder. 

This script can be executed at any time. But also in this repository is a .bat file which is set up to call get_info_from_inbox.py every 60 seconds. I have configured Windows Task Scheduler to execute the .bat file at each logon.

~~~~~
Email sender: "web.requests@adelaide.edu.au"
Email subject: "Webform submission - Knee function study"

Submitted on Mon, 12/16/2019 - 09:32
Submitted by: Anonymous
Submitted values are:
Full name
Test Subject

Age
100

Gender
Male

Phone number
000

Email address
email_address@gmail.com

Height
180

Weight
80
~~~~~
