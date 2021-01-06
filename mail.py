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
    def __init__(self, dataframe_user, email_sender='gemalabonf@outlook.fr', password='Test_pour_Appli'):
        self.email_sender = email_sender
        self.password = password
        self.dataframe_user = dataframe_user
        self.email_messages = self.email_infos()
        # a developper pour tirer les infos interessantes
    # send mail

    def email_infos(self):
        # récupérer les infos du dataframe et insérer les variables dans une liste de messages
        infos = self.dataframe_user
        email_messages = []
        for i in range(len(infos)):
            message = "Cher Monsieur, Madame %s, \
                \n\nNous nous permettons de vous relancer concernant l’arbre n°%d, %s, situé aux coordonnées [%f, %f].\
                \nNous sommes intervenus sur cet arbre le %s, pour %s.\
                \nNous souhaitons effectuer une nouvelle visite pour %s sur cet arbre dans les prochains mois.\
                \n\nNous vous prions de nous recontacter pour confirmer cette nouvelle action à effectuer au plus tard le %s.\
                \n\nBien cordialement,\n\nONF" % (
                    infos.iloc[i]['Nom du client'], 
                    infos.iloc[i]['Numéro de l\'arbre'], 
                    infos.iloc[i]["Nom français de l\'arbre"], 
                    infos.iloc[i]['x'], 
                    infos.iloc[i]['y'], 
                    infos.iloc[i]["Date du relevé"], 
                    "[type_operation]", # à modifier
                    "[type_operation]", # à modifier
                    "[date_limite]") # à modifier
            email_messages.append(message)
        return email_messages


    def send_emails(self, attachment_location = ''):
    
        #email_sender = 'gemalabonf@outlook.fr' # 'your_email_address@your_server.com'

        for email_message in self.email_messages:
    
            msg = MIMEMultipart()
            msg['From'] = self.email_sender # OK
            msg['To'] = self.email_sender # à modifier (cf ci dessous)
            # vérifier unicité mail dans cellule du dataframe
            # modifier self.email_sender en self.dataframe_user['Coordonnées mail']
            msg['Subject'] = "Relance pour %s" % ("[type_operation]") # à modifier
        
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
                server.login(self.email_sender, self.password) # OK
                text = msg.as_string()
                server.sendmail(self.email_sender, self.email_sender, text) # à modifier (cf ci dessous)
                # vérifier unicité mail dans cellule du dataframe
                # modifier self.email_sender en self.dataframe_user['Coordonnées mail']
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

        for email_message in self.email_messages:

            try:
                tls_context = ssl.create_default_context()
                server = imaplib.IMAP4('outlook.office365.com')
                server.starttls(ssl_context=tls_context)
                server.login(self.email_sender, self.password)
                # Select mailbox
                server.select("DRAFTS")
                # Create message
                new_message = email.message.Message()
                new_message["From"] = self.email_sender # OK
                new_message["To"] = self.email_sender # à modifier (cf ci dessous)
                # vérifier unicité mail dans cellule du dataframe
                # modifier self.email_sender en self.dataframe_user['Coordonnées mail']
                new_message["Subject"] = "Relance pour %s" % ("[type_operation]") # à modifier
                new_message.set_payload(email_message)
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