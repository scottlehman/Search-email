import win32com.client
import glob
import os

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
path = R'M:\Inbox'

for filename in glob.glob(os.path.join(path, '*.msg')):
    msg = outlook.OpenSharedItem(R'M:\Inbox\{}'.format(filename[58:]))
    print(msg.Sender)
    # print(msg.SenderEmailAddress)
    # print(msg.SentOn)
    # print(msg.To)
    # print(msg.CC)
    # print(msg.BCC)
    # print(msg.Subject)
    # print(msg.Body)

del outlook, msg
