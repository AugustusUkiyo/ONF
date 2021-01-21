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
from datetime import datetime
import os.path
import imaplib
import ssl
import email.message
import email.charset
import time

# gemalabonf@outlook.fr
# Test_pour_Appli

class mail:
    def __init__(self, dataframe, email_sender, password):
        self.email_sender = email_sender
        self.password = password
        self.dataframe = dataframe.sort_values(by=['Nom du client','Date de relance'])
        print(self.dataframe)
        self.recipients = self.dataframe['Nom du client'].unique() # A MODIFIER cf commentaires ci dessous
        # Récupérer les clients distincts dans dataframe (qui doit être ordonné dans dataframe.py)
        # pour chaque client récupérer la première ligne en brut dans un premier temps... puis améliorer
        print(self.recipients)

        self.email_messages = self.email_infos()
        print(self.email_messages)

    def email_infos(self):
        # récupérer les infos du dataframe et insérer les variables dans une liste de messages
        clients_infos = self.recipients # liste du nom des clients
        print(type(clients_infos))
        trees_infos = self.dataframe
        email_messages = []
        mois = ["None", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        for nom in clients_infos:
            print("nom", nom)
            # prendre la 1ere ligne
            print("trees", trees_infos)
            ligne_info_client = trees_infos[trees_infos['Nom du client']==nom]
            print("infos_client", ligne_info_client)
            print("date_relevé", ligne_info_client["Date du relevé"])
            print("type_date_relevé", type(ligne_info_client["Date du relevé"]))
            print("date_relevé_slice", ligne_info_client["Date du relevé"].iloc[0])
            print("type(date_relevé_slice):", type(ligne_info_client["Date du relevé"].iloc[0]))
            message = (
                "Madame, Monsieur " +  ligne_info_client["Nom du client"].iloc[0] + ", " +
                "\n\nNos équipes sont intervenues durant le mois de " +
                mois[ligne_info_client["Date du relevé"].iloc[0].date().month] + " " +
                str(ligne_info_client["Date du relevé"].iloc[0].date().year) + ", " +
                "pour réaliser un inventaire ainsi qu’un diagnostic visuel et sonore " +
                "du patrimoine arboré de votre site : " + ligne_info_client["Site"].iloc[0] + "." +
                "\n\nA l’issu de cette intervention, nous avons effectué des préconisations " +
                "de suivis et travaux pour les arbres qui présentaient des défauts en évolution." +
                "\n\nSont concernés par ces préconisations :"
                )

            ligne_info_client = trees_infos[trees_infos['Nom du client']==nom]
            print(message)
            print("deadline =", ligne_info_client["Deadline"].iloc[0])
            print("opération =", ligne_info_client["Type d'opération"].iloc[0])

            for j in range(len(ligne_info_client)):
                print(j)
                #print(clients_infos.iloc[i]["Type d’opération"])
                print(ligne_info_client["Type d'opération"].iloc[j])
                addendum = (
                    "\n\n- Arbre numéro " + str(ligne_info_client['Numéro de l\'arbre'].iloc[j]) +
                    ", " + ligne_info_client["Nom français de l\'arbre"].iloc[j] + " " +
                    "situé aux coordonnées [" + str(ligne_info_client['x'].iloc[j]) + ", " +
                    str(ligne_info_client['y'].iloc[j]) + "] :" +
                    "\nPréconisation(s): " + ligne_info_client["Type d'opération"].iloc[j]
                    )

                message = message + addendum
            
            message = message + (
                "\n\nParce que valoriser et préserver votre patrimoine arboré, " +
                "c’est agir sur la protection de vos biens et la sécurité des riverains, " + 
                "nous vous invitons à recontacter votre interlocuteur ONF avant le " +
                str(ligne_info_client['Deadline'].iloc[0]) + 
                " si vous souhaitez effectuer un contrôle de ces arbres." +
                "\n\n\nL’équipe ONF"
                )

            print("final_message =", message)
                
            email_messages.append([nom, ligne_info_client["Site"].iloc[0], message, ligne_info_client["Coordonnées mail"].iloc[0]])

        return email_messages


    def send_emails(self, attachment_location = ''):
    
        #email_sender = 'gemalabonf@outlook.fr' # 'your_email_address@your_server.com'
        list_msg = ''

        for i in range(len(self.email_messages)):
    
            msg = MIMEMultipart()
            msg['From'] = self.email_sender # OK
            msg['To'] = self.email_messages[i][3] # à modifier (cf ci dessous)
            # vérifier unicité mail dans cellule du dataframe
            # modifier self.email_sender en self.recipients.iloc[i]['Coordonnées mail']
            msg['Subject'] = "ONF : pensez à réaliser le suivi de votre patrimoine arboré %s" % (
                self.email_messages[i][1])
        
            msg.attach(MIMEText(self.email_messages[i][2], 'plain'))
        
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
                server.sendmail(self.email_sender, self.email_messages[i][3], text) # à modifier (cf ci dessous)
                # vérifier unicité mail dans cellule du dataframe
                # modifier self.email_sender en self.recipients.iloc[i]['Coordonnées mail']
                print('email sent')
                server.quit()
                list_msg += "Le mail a bien été envoyé à "+ self.email_messages[i][0] + "\n"
            except:
                print("SMPT server connection error")
                list_msg += "Il y a un problem d'envoyer le mail à " + self.email_messages[i][0] + "\n"

        return list_msg

    # Stocker mail comme brouillon
    def create_draft(self):
        
        print(len(self.email_messages))
        for i in range(len(self.email_messages)):
            #print("site = ", self.recipients["Site"].iloc[i])
            print(i)

            try:
                #print("site = ", self.recipients["Site"].iloc[i])
                tls_context = ssl.create_default_context()
                server = imaplib.IMAP4('outlook.office365.com')
                server.starttls(ssl_context=tls_context)
                server.login(self.email_sender, self.password)
                # Select mailbox
                server.select("DRAFTS")
                # Create message
                new_message = email.message.Message()
                new_message["From"] = self.email_sender # OK
                new_message["To"] = self.email_messages[i][3] # à modifier (cf ci dessous)
                # vérifier unicité mail dans cellule du dataframe
                # modifier self.email_sender en self.recipients.iloc[i]['Coordonnées mail']
                new_message["Subject"] = "ONF : pensez à réaliser le suivi de votre patrimoine arboré %s" % (
                    self.email_messages[i][1])
                new_message.set_payload(self.email_messages[i][2])
                # Fix special characters by setting the same encoding we'll use later to encode the message
                new_message.set_charset(email.charset.Charset("utf-8"))
                encoded_message = str(new_message).encode("utf-8")
                server.append('DRAFTS', '', imaplib.Time2Internaldate(time.time()), encoded_message)
                # Cleanup
                server.close()
            except:
                print("Server connection error")

        return True