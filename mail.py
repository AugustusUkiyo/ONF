#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Dec  9 12:45:38 2020

@author: ukiyo
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path
import imaplib
import ssl
import email.message
import email.charset
import time


class mail:
    def __init__(self, email_sender='gemalabonf@outlook.fr', password='Test_pour_Appli', dataframe_user):
        self.email_sender = email_sender
        self.password = password
        self.dataframe_user =dataframe_user
        # a developper pour tirer les infos interessantes
    # send mail
    def send_email(self, email_recipient, 
                   email_subject,
                   email_message,
                   attachment_location = ''):
    
        #email_sender = 'gemalabonf@outlook.fr' # 'your_email_address@your_server.com'
    
        msg = MIMEMultipart()
        msg['From'] = self.email_sender
        msg['To'] = email_recipient
        msg['Subject'] = email_subject
    
        msg.attach(MIMEText(email_message, 'plain'))
    
        if attachment_location != '':
            filename = os.path.basename(attachment_location)
            attachment = open(attachment_location, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            "attachment; filename= %s" % filename)
            msg.attach(part)
    
        try:
            server = smtplib.SMTP('smtp.office365.com', 587)
            server.ehlo()
            server.starttls()
            server.login(self.email_sender, self.password)
            text = msg.as_string()
            server.sendmail(self.email_sender, email_recipient, text)
            print('email sent')
            server.quit()
        except:
            print("SMPT server connection error")
            return False
        return True
    """
    send_email('gemalabonf@outlook.fr',
               'Happy New Year',
               'We love Outlook', 
               '')
    """
    # Stocker mail comme brouillon
    def create_draft(self):
        try:
            tls_context = ssl.create_default_context()
            server = imaplib.IMAP4('outlook.office365.com')
            server.starttls(ssl_context=tls_context)
            server.login(self.email_sender, self.password)
            # Select mailbox
            server.select("DRAFTS")
            # Create message
            new_message = email.message.Message()
            new_message["From"] = "Sven LOTHE <gemalabonf@outlook.fr>"
            new_message["To"] = "Sven Lothe <gemalabonf@outlook.fr>"
            new_message["Subject"] = "Your subject"
            new_message.set_payload("""
            This is your message.
            It can have multiple lines and
            contain special characters: äöü.
            """)
            # Fix special characters by setting the same encoding we'll use later to encode the message
            new_message.set_charset(email.charset.Charset("utf-8"))
            encoded_message = str(new_message).encode("utf-8")
            server.append('DRAFTS', '', imaplib.Time2Internaldate(time.time()), encoded_message)
            # Cleanup
            server.close()
        except:
            print("Server connection error")
            return False
        return True