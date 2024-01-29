import os
import shutil
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import win32com.client as win32


# folder path
dir_path =os.getcwd()
email_path=dir_path+"//email"
zeugnisse_path=dir_path+"//zeugnisse"

# list to store files
res = []
res2= []

# Iterate directory
for path in os.listdir(email_path):
    # check if current path is a file
    if os.path.isfile(os.path.join(email_path, path)):
        res2.append(path)

# Iterate directory
for path in os.listdir(zeugnisse_path):
    # check if current path is a file
    if os.path.isfile(os.path.join(zeugnisse_path, path)):
        res.append(path)
print(res)
print(res2)

for x in res:
  print(x)
  if x in res2: 
    print("exist") 
  else: 
    print("not exist")
    
    olApp = win32.Dispatch('Microsoft Outlook.Application')
    olNS = olApp.GetNameSpace('MAPI')
    mailItem=olApp.CreateItem(0)
    mailItem.Subject='dummy'
    mailItem,BodyFormat=1
    mailItem.Body="hello"
    mailItem.To='jonczykjakub@outllook.com'
    mailItem.Attachments.Add(zeugnisse_path + x)
    mailItem.Display()
   # olmailitem=0x0 #size of the new email
   # newmail=ol.CreateItem(olmailitem)
   # newmail.Subject= 'Testing Mail'
   # newmail.To='jonczykjakub@outllook.com'
    #newmail.Body= 'Hello, this is a test email.'

    #attach=zeugnisse_path + x
    #newmail.Attachments.Add(attach)

    # To display the mail before sending it
   # newmail.Display() 

    # newmail.Send()


# Ending messagebox
messagebox.showinfo(title="Message", message="Files have been copied")
