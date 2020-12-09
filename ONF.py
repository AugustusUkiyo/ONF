#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Dec  7 15:01:20 2020

@author: ukiyo
"""

# Constants
GREEN = "#9bdeac"
YELLOW = "#f7f5dd"

# Library
import pandas as pd
import numpy as np
from datetime import datetime
from tkinter import *
from tkinter import messagebox
import smtplib

df = pd.read_excel('Excel_type.xlsx', sheet_name='DVS_v2_0')
df_final = df[['GlobalID', 'Nom du client', 'Coordonnées mail', 'Numéro de l\'arbre', 
               "Date du relevé", "Nom français de l\'arbre", 'Type de contrôle/suivi', 
               'Date de relance pour contrôle/suivi', "Type d'intervention 1", 
               'Date de relance pour intervention 1', "Type d'intervention 2", 
               'Date de relance pour intervention 2', 'x', 'y']]
date_today = datetime.today().date()
date_today_str = datetime.today().strftime('%m/%d/%Y')

# Convert
date1 = [ x.date() for x in list(df_final['Date de relance pour contrôle/suivi'])]
date2 = [ x.date() for x in list(df_final['Date de relance pour intervention 1'])]
date3 = [ x.date() for x in list(df_final['Date de relance pour intervention 2'])]

# Verifier les dates de relances
bool_date1 = np.array([ date_today > y for y in date1])
bool_date2 = np.array([ date_today > y for y in date2])
bool_date3 = np.array([ date_today > y for y in date3])

# Ajouter dans la colonne de dataframe
bool_relance = bool_date1 + bool_date2 + bool_date3
df_final['Relance mail ce jour '+date_today_str] = bool_relance

# Retourner tous les infos du clients aui ont le date de relance aujourd'hui
df_return = df_final[df_final['Relance mail ce jour '+date_today_str]==True]

# Verify user
def verify_user():
    """
    Retourner le nombre de personne dont on va envoyer le mail
    """
    Excel = source_excel_entry.get()
    # Messagebox
    msg = "Il y a "+str(len(df_return))+" personnes qui ont la date de relancement aujourd'hui"
    messagebox.showinfo(title="Information", message=msg)

# Create window
window = Tk()
window.title("ONF")
window.config(padx=20, pady=20)

# Canvas
canvas = Canvas(width=200, height=200)
logo_img = PhotoImage(file="logo.png")
canvas.create_image(100,100,image=logo_img)
canvas.grid(row=0,column=1)

# Labels
source_excel = Label(text="Fichier Excel")
source_excel.grid(row=1, column=0)
email_label = Label(text="Email")
email_label.grid(row=2, column=0)
password_label = Label(text="Password")
password_label.grid(row=3, column=0)

# Entries
source_excel_entry = Entry(width=35)
source_excel_entry.grid(row=1, column=1)
email_entry = Entry(width=35)
email_entry.grid(row=2,column=1)
password_entry = Entry(width=35, show='*')
password_entry.grid(row=3, column=1)

# Button
# Open file excel
open_file_button = Button(text="Open file Excel")
open_file_button.grid(row=1, column=3)
# Verify contenu file Excel
verify_button = Button(text="Verify User",width=15,command=verify_user)
verify_button.grid(row=4, column=1)
# Create crafts
produce_mail_button = Button(text="Produce mail",width=15)
produce_mail_button.grid(row=5, column=1)
# Sent mail
sent_mail_button = Button(text="Sent mail",width=15)
sent_mail_button.grid(row=6, column=1)

window.mainloop()

# send mail
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path


def send_email(email_recipient,
               email_subject,
               email_message,
               attachment_location = ''):

    email_sender = 'gemalabonf@outlook.fr' # 'your_email_address@your_server.com'

    msg = MIMEMultipart()
    msg['From'] = email_sender
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
        server.login(email_recipient, 'Test_pour_Appli')
        text = msg.as_string()
        server.sendmail(email_sender, email_recipient, text)
        print('email sent')
        server.quit()
    except:
        print("SMPT server connection error")
    return True
"""
send_email('gemalabonf@outlook.fr',
           'Happy New Year',
           'We love Outlook', 
           '')
"""
# Stocker mail comme brouillon
import imaplib
import ssl
import email.message
import email.charset
import time
def create_draft():
    tls_context = ssl.create_default_context()
    server = imaplib.IMAP4('outlook.office365.com')
    server.starttls(ssl_context=tls_context)
    server.login('gemalabonf@outlook.fr', 'Test_pour_Appli')
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























