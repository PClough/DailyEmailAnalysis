# -*- coding: utf-8 -*-


import win32com.client
import win32com
import datetime
import pytz
import pandas as pd

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts;
Todays_Date = datetime.datetime.now(pytz.UTC)

#Folder_list_to_search = [0,1,3]
# =============================================================================
# Order of folders in outlook
#     Deleted Items
#     Inbox
#     Outbox
#     Sent Items
#     Junk Email
#     Files
#     Scheduled
#     RSS Subscriptions
#     Calendar
#     Yammer Root
#     PersonMetadata
#     ExternalContacts
#     Archive
#     Working Set
#     Tasks
#     Conversation Action Settings
#     Conversation History
#     Drafts
#     Journal
#     Contacts
#     Quick Step Settings
#     Sync Issues
#     Notes
#     Social Activity Notifications
# =============================================================================


for account in accounts:
    myoutlook = outlook.Folders(account.DeliveryStore.DisplayName).Folders
    
    ## Inbox
    myinbox = myoutlook[1].Items
    n = 0
    my_emails_Sender = list()
    for i in myinbox:
        try:
            my_emails_Sender.append(i.Sender())
        except:
            n = n + 1
            # These are usually calendar events in the inbox
            #print("Cheese? " + str(i))
            #print("Hate self... Why Cheesoid exist?")
    
    freq_Senders = pd.Series(my_emails_Sender).value_counts().reset_index()
    
    
    ## Sent box
    mySentbox = myoutlook[3].Items
    n = 0
    my_emails_Sent = list()
    for i in mySentbox:
        try:
            recipients = i.Recipients
            for r in recipients:
                my_emails_Sent.append(str(r))
        except:
            n = n + 1
            # These are usually calendar events in the inbox
            #print("Cheese? " + str(i))
            #print("Hate self... Why Cheesoid exist?")
    
    freq_Sent = pd.Series(my_emails_Sent).value_counts().reset_index()
     
    
    ## Deleted box
    myDeletedbox = myoutlook[0].Items
    n = 0
    my_emails_Deleted = list()
    for i in myDeletedbox:
        try:
            my_emails_Deleted.append(i.Sender())
        except:
            n = n + 1
            # These are usually calendar events in the inbox
            #print("Cheese? " + str(i))
            #print("Hate self... Why Cheesoid exist?")
    
    freq_Deleted = pd.Series(my_emails_Deleted).value_counts().reset_index()


print("Finished Succesfully")


# =============================================================================
# To Do: 
# box around plot
# =============================================================================

# https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._mailitem?view=outlook-pia
# https://stackoverflow.com/questions/22813814/clearly-documented-reading-of-emails-functionality-with-python-win32com-outlook

#The objects used above have the following functionality:
#inbox -
#.Folders
#.Items

#messages - https://docs.microsoft.com/en-us/office/vba/api/outlook.items
#.GetFirst()
#.GetLast()
#.GetNext()
#.GetPrevious()
#.Attachments()

#message -
#.Subject
#.Body
#.To
#.Recipients
#.Sender()
#.Sender.Address
#.ReceivedTime
#.CreationTime
#.SentOn
#.Attachments
#.Sent

#attachments -
#.item()
#.Count()

#attachment -
#.filename()
