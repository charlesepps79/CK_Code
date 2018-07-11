# -*- coding: utf-8 -*-
"""
Created on Wed Jul 11 09:46:17 2018

@author: cepps
"""

import imaplib, email, os
import configparser

config = configparser.ConfigParser()
config.read(r"E:\cepps\Web_Report\Credit_Karma\etc\config.txt")
password_config = config.get("configuration","password")

user = 'cepps@regionalmanagement.com'
password = password_config
imap_url = 'imap-mail.outlook.com'

#Where you want your attachments to be saved (ensure this directory exists) 
attachment_dir = r'E:\cepps\Web_Report\Credit_Karma\attachments'

# sets up the auth
def auth(user,password,imap_url):
    con = imaplib.IMAP4_SSL(imap_url)
    con.login(user,password)
    return con

# extracts the body from the email
def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)
    
# allows you to download attachments
def get_attachments(msg):
    for part in msg.walk():
        if part.get_content_maintype()=='multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()

        if bool(fileName):
            filePath = os.path.join(attachment_dir, fileName)
            with open(filePath,'wb') as f:
                f.write(part.get_payload(decode=True))
                
#search for a particular email
def search(key,value,con):
    result, data  = con.search(None,key,'"{}"'.format(value))
    return data

#extracts emails from byte array
def get_emails(result_bytes):
    msgs = []
    for num in result_bytes[0].split():
        typ, data = con.fetch(num, '(RFC822)')
        msgs.append(data)
    return msgs

con = auth(user,password,imap_url)
con.select('INBOX/CK_Reports')

result, data = con.fetch(b'10','(RFC822)')
raw = email.message_from_bytes(data[0][1])
get_attachments(raw)

