# -*- coding: utf-8 -*-


import win32com.client
import win32com
import pandas as pd
from textblob import TextBlob
import matplotlib.pyplot as plt

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts;

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


#%% Sentiment analysis
# it returns a tuple representing polarity and subjectivity of each tweet. 
# Here, we only extract polarity as it indicates the sentiment as value nearer to 1 means a positive sentiment
# and values nearer to -1 means a negative sentiment. 

for account in accounts:
    myoutlook = outlook.Folders(account.DeliveryStore.DisplayName).Folders
    
    #%% Inbox
    myinbox = myoutlook[1].Items
    n = 0
    my_emails_time = list()
    email_Polarity = list()
    email_Subjectivity = list()
    for i in myinbox:
        try:
            my_emails_time.append(i.ReceivedTime.date())
            email_Polarity.append(TextBlob(i.body).sentiment[0])
            email_Subjectivity.append(TextBlob(i.body).sentiment[1])
        except:
            n = n + 1
    
    my_Inbox_sentiment = pd.DataFrame({'Dates' : my_emails_time,
                                        'Polarity' : email_Polarity,
                                        'Subjectivity' : email_Subjectivity})
    
    plt.plot(my_Inbox_sentiment.Dates, my_Inbox_sentiment.Polarity)
    plt.xlabel('Time / s')
    plt.ylabel('Sentiment')
    plt.title('Received')
    plt.show()
    
    
    #%% Sent box
    mySentbox = myoutlook[3].Items
    n = 0
    my_emails_time = list()
    email_Polarity = list()
    email_Subjectivity = list()
    for i in mySentbox:
        try:
            my_emails_time.append(i.ReceivedTime.date())
            email_Polarity.append(TextBlob(i.body).sentiment[0])
            email_Subjectivity.append(TextBlob(i.body).sentiment[1])
        except:
            n = n + 1
           
    my_sentbox_sentiment = pd.DataFrame({'Dates' : my_emails_time,
                                        'Polarity' : email_Polarity,
                                        'Subjectivity' : email_Subjectivity})
    
    plt.plot(my_sentbox_sentiment.Dates, my_sentbox_sentiment.Polarity)
    plt.xlabel('Time / s')
    plt.ylabel('Sentiment')
    plt.title('Sent')
    plt.show()
    
    #%% Deleted box
    myDeletedbox = myoutlook[0].Items
    n = 0
    my_emails_time = list()
    email_Polarity = list()
    email_Subjectivity = list()
    for i in myDeletedbox:
        try:
            my_emails_time.append(i.ReceivedTime.date())
            email_Polarity.append(TextBlob(i.body).sentiment[0])
            email_Subjectivity.append(TextBlob(i.body).sentiment[1])
        except:
            n = n + 1
           
    my_deletedbox_sentiment = pd.DataFrame({'Dates' : my_emails_time,
                                        'Polarity' : email_Polarity,
                                        'Subjectivity' : email_Subjectivity})
    
    plt.plot(my_deletedbox_sentiment.Dates, my_deletedbox_sentiment.Polarity)
    plt.xlabel('Time / s')
    plt.ylabel('Sentiment')
    plt.title('Deleted')
    plt.show()
    
    
print("Finished Succesfully")


#%%
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
#.Attachments

#message -
#.Subject
#.Body
#.To
#.Recipients
#.Sender
#.Sender.Address
#.ReceivedTime
#.CreationTime
#.SentOn
#.Attachments
#.Sent

#attachments -
#.item()
#.Count

#attachment -
#.filename
