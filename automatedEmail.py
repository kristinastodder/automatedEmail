#! python3
# Automated Project Email Reminder.py - Sends emails based on yes or no status in Excel spreadsheet.

#Key: FROM EMAIL HERE - the email that you want to send from, CC EMAIL HERE - the email that you want to CC

import openpyxl, smtplib, sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Open the spreadsheet and get the latest status.
wb = openpyxl.load_workbook('Email Program.xlsx')
sheet = wb['Sheet1']

# Check each person's status.
receiveEmail = {}
for r in range(2, sheet.max_row + 1):
    sendIt = sheet.cell(row=r, column=1).value
    if sendIt != 'no':
        name = sheet.cell(row=r, column=3).value
        email = sheet.cell(row=r, column=6).value
        receiveEmail[name] = email

# Log in to email account.
smtpObj = smtplib.SMTP('smtp-mail.outlook.com', 587)
smtpObj.ehlo()
smtpObj.starttls()
password = input('Enter your password:')
smtpObj.login('FROM EMAIL HERE', password)


# Send out reminder emails.
for name, email in receiveEmail.items():
    from_addr = 'FROM EMAIL HERE'
    to_addr = email
    cc_addr = 'CC EMAIL HERE'
    
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Cc'] = cc_addr
    msg['Subject'] = 'New Project'
    body = "Hi %s,\n\n\
We have a new project called NAME HERE. I invited you through Building Connected and I want to make sure that you received the invite.\n\n\
Please let me know if you are interested in bidding this one for us.\n\n\
It is due next Tuesday. \n\n\
Thank you,\n\n\
Kristina Stodder\n\
Estimating Coordinator\n\
COMPANY NAME\n\
Direct: PHONE NUMBER HERE\n\
Estimating Fax: FAX NUMBER HERE" %(name)
    msg.attach(MIMEText(body, 'plain'))
    text = msg.as_string()
                    
    sendmailStatus = smtpObj.sendmail(from_addr, to_addr, text)

if sendmailStatus != {}:
        print('There was a problem sending email to %s: %s' % (email, sendmailStatus))

# Cc CC EMAIL HERE
for name, email in receiveEmail.items():
    from_addr = 'FROM EMAIL HERE'
    to_addr = 'CC EMAIL HERE'
    cc_addr = email
    
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = cc_addr
    msg['Cc'] = to_addr
    msg['Subject'] = 'New Project'
    body = "Hi %s,\n\n\
We have a new project called PROJECT NAME HERE. I invited you through Building Connected and I want to make sure that you received the invite.\n\n\
Please let me know if you are interested in bidding this one for us.\n\n\
It is due next Tuesday. \n\n\
Thank you,\n\n\
Kristina Stodder\n\
Estimating Coordinator\n\
COMPANY NAME HERE\n\
Direct: PHONE NUMBER HERE\n\
Estimating Fax: FAX NUMBER HERE" %(name)
    msg.attach(MIMEText(body, 'plain'))
    text = msg.as_string()
    
    print('Sending email to %s at %s and copying CC EMAIL HERE' %(name, email))
                    
    sendmailStatus = smtpObj.sendmail(from_addr, to_addr, text)

if sendmailStatus != {}:
        print('There was a problem sending email to %s: %s' % (email, sendmailStatus))

smtpObj.quit()

print('Complete')

