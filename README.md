# mailbox-autoMonitor
Console program to automatically monitor emails leveraging Microsoft Graph API

# pre-requisites
pip install requirements.txt /pip3 install requirements.txt (python3)

# usage
Register an application in Microsoft Azure active directory. For detailed steps refer : (https://docs.microsoft.com/en-us/azure/active-directory/develop/howto-create-service-principal-portal)

Configure the config1.json with the client id, tenant id obtained after creating the application in Azure AD.

Note: you can use multiple configuration files to track multiple mailboxes

## Read Inbox Mails
Directly run the application by python3 monitor.py in terminal or python monitor.py in cmd to read the latest unread mail from the inbox
Note: the script marks the mail as read after reading the email.

run the monitor.py in endless while to continuesly monitor the emails, this can be used to trigger other applications or scripts through an API call.
```
my_mailbox = MailboxMonitor("config1.json")
while True:
  try:
    recent_unread_mail = my_mailbox.read_inbox_mails()
    #call an api or perform further steps
  except Exception as ex:
    logging.error(str(ex))
```

## Read mails from inbox Child Folders
use the method read_folder_mails to read the latest unread mails from child folders.
use a for-loop over all the child folders in a list or maintain a seperate file to add the additional child folders
Iterate through all the child folders and read the latest mail.
```
my_mailbox = MailboxMonitor("config1.json")
child_folders = ['folder1','folder2','folder3']
for folder in child_folders:
  recent_unread_mail = my_mailbox.read_folder_mails(folder)
  #call an api or perform further steps
```

## Read mails from sent items
Calling the method read_sent_mails fetches the last 5 emails from the sent items folder
```
my_mailbox = MailboxMonitor("config1.json")
while True:
  try:
    recent_unread_mail = my_mailbox.read_sent_mails()
    #call an api or perform further steps
  except Exception as ex:
    logging.error(str(ex))
```

## Download attachment from email
```
my_mailbox = MailboxMonitor("config1.json")
while True:
  try:
    my_mailbox.get_attachment_in_file(message_id,file)
  except Exception as ex:
    logging.error(str(ex))
```
