import datetime
import win32com.client
from os import path
import csv, json

local_path = path.expanduser(r'\\fsnas100\...') #your file path goes here
now = datetime.date.today()

# connect to Outlook email
Outlook = win32com.client.Dispatch('Outlook.Application')
mapi = Outlook.GetNamespace('MAPI')
inbox = mapi.GetDefaultFolder(6)

#insert name of folder where emails stored below (not necessary)
ga_data_folder = inbox.Folders.Item('GA Daily Reports') 

messages = ga_data_folder.Items

#specific subject line 
message_subject = 'FW: Google Analytics: Daily'

which_item = ''
which_item = str(which_item)

def grab_attachments(message_subject,which_item):
    for message in messages:
        if message.Subject == message_subject and message.Unread:
            attachments = message.Attachments
            #print(attachments)
            message.Unread = False

            for attach in attachments:
                #print(attach)
                attach.SaveAsFile(path.join(local_path,str(attach)))
                break
            
            # convert to json and save to SQL Server
            def to_json(attach):
                print(attach)
                    
            to_json(attach)

grab_attachments(message_subject,which_item)
