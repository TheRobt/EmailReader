import win32com.client, sqlite3
from datetime import datetime
import time
from inspect import getmembers
import os

path = os.path.expanduser("~/Desktop/Attachments")
def collectMail():
    conn = sqlite3.connect('outlook.db')
    i = 0
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        print('total messages: ', len(messages))
        message = messages.GetLast()
        print(message)
        while message:
            i += 1
            subject = message.Subject
            print(i, subject)
            received_time = str(message.ReceivedTime)
            print(received_time)
            #received_time = datetime.strptime(received_time, '%m/%d/%y %H:%M:%S')
            print (message.EntryID)
            html_body = message.Body
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment)))
            size = message.Size
            sender =  message.SenderName
            receiver = message.To
            cc = message.Cc
            body = message.Body
                #conn.execute('insert into outlook(SUBJECT, SENDER, RECEIVER, CC, SIZE, RECEIVED_TIME, BODY, HTML_BODY) values(?, ?, ?, ?, ?, ?, ?, ?)', (subject, sender, receiver, cc, size, received_time, body, html_body))
                #conn.commit()
            message = messages.GetPrevious()
            message = 0
            time.sleep(1)
    finally:
        print('connection closed')
        conn.close()
collectMail()
