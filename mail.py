import win32com.client
import os
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

for account in mapi.Accounts:
    print(account.DeliveryStore.DisplayName)

#To Access Inbox
inbox = mapi.GetDefaultFolder(6)

#To Access Folders
inbox = mapi.GetDefaultFolder(6).Folders["Folder Name"]
