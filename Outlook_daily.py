# -*- coding: utf-8 -*-

import win32com.client
import win32com
import datetime
import chart_studio
import plotly.graph_objects as go
from plotly.offline import plot

chart_studio.tools.set_credentials_file(username='YOUR_UN', api_key='YOUR_API_CODE')

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts;
Todays_Date = datetime.date.today()

Folder_list_to_search = [0,1,3]
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

#%%
for account in accounts:
    y_Data_deleted = 0
    y_Data_inbox = 0
    y_Data_sent = 0
    for k in Folder_list_to_search:
        myoutlook = outlook.Folders(account.DeliveryStore.DisplayName).Folders
        my_emails = myoutlook[k].Items
        for i in my_emails:
            try:
               if i.ReceivedTime.date() == Todays_Date:
                    if k == 0:
                        y_Data_deleted = y_Data_deleted + 1
                    elif k == 1:
                        y_Data_inbox = y_Data_inbox + 1
                    elif k == 3:
                        y_Data_sent = y_Data_sent + 1
            except:
                 continue

# Total emails
total_emails = y_Data_deleted + y_Data_inbox + y_Data_sent


#%% Plot data
# Deleted emails
emails_deleted = go.Bar(
            x = [Todays_Date],
            y = [y_Data_deleted],
            name = "Emails deleted",
            marker = dict(color = '#f22121')
)
# Inbox emails
emails_inbox = go.Bar(
            x = [Todays_Date],
            y = [y_Data_inbox],
            name = "Emails received",
            marker = dict(color = '#f48642')
)
# Sent emails
emails_sent = go.Bar(
            x = [Todays_Date],
            y = [y_Data_sent],
            name = "Emails sent",
            marker = dict(color = '#4286f4')
)
# Total emails
emails_total = go.Bar(
            x = [Todays_Date],
            y = [total_emails],
            name = "Total emails sent and received",
            marker = dict(color = '#42b9bf')
) 

# https://images.plot.ly/plotly-documentation/images/python_cheat_sheet.pdf
layout = dict(
    showlegend = True,
    autosize = False,
    width = 900,
    height = 900,
    margin=go.layout.Margin(
        l = 50,
        r = 50,
        b = 100,
        t = 100,
        pad = 4
    ),

    yaxis = dict(
        title = 'Number of emails',
        showgrid = False,
        showline = True,
        ticks = 'inside'
    ),
    
    xaxis = dict(
        title = 'Time',
        showgrid = False,
        showline = True,
        ticks = 'inside',
        type = 'date'
    )
)

fig = dict(data = [emails_deleted, emails_inbox, emails_sent, emails_total], layout = layout)
plot(fig)


print("Finished Succesfully")


#%% =============================================================================

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
