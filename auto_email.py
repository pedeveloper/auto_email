#!/usr/bin/python3
#Requires pip package: pywin32
#Currently only works for Windows OS running Outlook

import smtplib
import os
import win32com.client as client
import codecs
from pywintypes import com_error
import csv

try:
    your_dir_name = input('Type the name of the folder containing the reports divided by recipient (Ex. mail_demo): ')
except KeyboardInterrupt: #Handles KeyboardInterrupt error to close the program
    print('Program closed due to keyboard interrupt.')
    quit()

working_dir = os.popen('echo %cd%').read() #Pull the current directory path. Needed because your_dir_name should be on the same root
recipients_sent_list = []
recipients_not_sent_list = []

#Change \ chars to / from working dir path so it can be accepted by the Outlook object later on
working_dir_rep1 = working_dir.replace(":\\","://")
working_dir_rep2 = working_dir_rep1.replace('\\','/')
working_dir_rep3 = working_dir_rep2.replace('\n','')

#save the current dir path
folders_string = os.popen('dir ' + your_dir_name + ' /B /AD').read()
folder_list = list(folders_string.split("\n"))

#Asks questions for important variables to be used in the email. Summary and Mitigation plan must be pasted as a one liner and in HTML format.
try:
    sender = input('Who is sending this email?: ')
    print('\n')
    ticket = input('Ticket number (ex. case_123): ')
    print('\n')
    title = input('Title: ')
    print('\n')
    summary = input('Summary: ')
    print('\n')    
    report_date = input('When were the reports pulled? (ex. 2022-02-15): ')
    print('\n')
    signature_name = input('What is your Outlook signature name? (ex. Work): ')    
    print('\n')
except KeyboardInterrupt: #Handles KeyboardInterrupt error to close the program
    print('Program closed due to keyboard interrupt.')
    quit()

#Create a list of the recipients so it can be later iterated
while("" in folder_list):
    folder_list.remove("")

def send_email(sender, recipient, report_date, recipient, ticket, title, summary, signature_name):

    outlook = client.Dispatch('Outlook.Application')
    message = outlook.CreateItem(0)    

    message.Subject = ticket + ' | ' + recipient + ' | ' + title
    message.CC = 'example1@example.com'

    signature = get_signature(signature_name)
    message.HTMLBody = '<div>Hello,<br><br>' + 'Summary' + summary + '<br><br>Thank you,<br><br>' + signature +'</div>' 
    message.SentOnBehalfOfName = sender
    
    file_path = "%s/%s/%s/%s_%s.xlsx" % (working_dir_rep3,your_dir_name,recipient,recipient,report_date)
    
    message.Attachments.Add(file_path)
    
    message.Save()    
    message.Send()
    

#Pulls the signature from your AppData Windows folder given the name specified in a previous input. Images do not get pulled correctly.
def get_signature(signature_name):
    signature_path = os.path.join((os.environ['USERPROFILE']),'AppData\Roaming\Microsoft\Signatures\%s_files\\' % (signature_name)) # Finds the path to Outlook signature files with signature name provided in the input
    html_doc = os.path.join((os.environ['USERPROFILE']),'AppData\Roaming\Microsoft\Signatures\%s.htm' % (signature_name))     #Specifies the name of the HTML version of the stored signature
    html_doc = html_doc.replace('\\\\', '\\') #Removes escape backslashes from path string

    html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore') #Opens HTML file and ignores errors
    signature_code = html_file.read()               #Writes contents of HTML signature file to a string
    signature_code = signature_code.replace('Work_files/', signature_path)      #Replaces local directory with full directory path    
    html_file.close()
    return signature_code

def print_recipients_sent_list():
    print("======")
    print("List of emails SENT✅: ")
    print("======")
    for recipient in range(len(recipients_sent_list)):
        print(recipients_sent_list[recipient])


def print_recipients_not_sent_list():
    print("======")
    print("List of emails NOT SENT❌: ")
    print("======")
    for recipient in range(len(recipients_not_sent_list)):        
        print(recipients_not_sent_list[recipient])


for recipient in folder_list:#Loop to iterate over the list that contains the recipient names.
    try:        
        send_email(sender,recipient,report_date,recipient,ticket,title,summary,signature_name)
    except com_error as e:
        if e.excepinfo[5] == -2147467259: #Handles error whenever an recipient does not match a valid DL. Prompts once for manual correction.
            print('Outlook failed to recognize one or more email recipients: ' + recipient)
            new_recipient = input('Type a new corresponding DL for recipient ' + recipient + ': ')
            if new_recipient=="":#skips to next recipient if we digit a blank string in 
                print("Blank string submitted. No email sent. Skipping to next recipient.\n")
                recipients_not_sent_list.append(recipient)
                continue
            try:
                send_email(sender,recipient,report_date,new_recipient,ticket,title,summary,signature_name)                
            except Exception as e:
                print('Oops, there was an error with the new recipient name. Mail not sent.')
                print(e)
                recipients_not_sent_list.append(recipient)             
    except KeyboardInterrupt: #Handles KeyboardInterrupt error to allow
        print('Program closed due to keyboard interrupt.')
    else:
        recipients_sent_list.append(recipient)
    finally:
        print('\n')

print("\n\n")
print_recipients_sent_list()
print("\n\n")
print_recipients_not_sent_list()