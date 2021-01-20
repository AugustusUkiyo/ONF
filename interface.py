#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Dec  9 12:45:40 2020

@author: ukiyo
"""


from tkinter import *
from tkinter import font
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
import dataframe
import mail
import pandas as pd 
from datetime import datetime
from datetime import timedelta
#from tkFileDialog import askdirectory


def get_date():
    """
    """
    if combo.get() == 'Dernière relance':

        fileHandle = open ( 'db.txt',"r" )
        lineList = fileHandle.readlines()
        fileHandle.close()
        date_debut = (datetime.strptime(lineList[len(lineList)-1][:-1], '%m/%d/%Y')).strftime('%m/%d/%Y')
    
    else:
        
        if combo.get() =="1 semaine":
            date_debut = (datetime.today().date() - timedelta(days=7)).strftime('%m/%d/%Y')
        if combo.get() =="2 semaines":
            date_debut = (datetime.today().date() - timedelta(days=14)).strftime('%m/%d/%Y')
        if combo.get() =="3 semaines":
            date_debut = (datetime.today().date() - timedelta(days=21)).strftime('%m/%d/%Y')
        if combo.get() =="1 mois":
            date_debut = (datetime.today().date() - timedelta(days=30)).strftime('%m/%d/%Y')
        if combo.get() =="2 mois":
            date_debut = (datetime.today().date() - timedelta(days=60)).strftime('%m/%d/%Y')
        if combo.get() =="3 mois":
            date_debut = (datetime.today().date() - timedelta(days=90)).strftime('%m/%d/%Y')
    
    date_fin = datetime.today().date().strftime('%m/%d/%Y')

    return date_debut, date_fin      
    

#nb_pers = 2
#open file excel
def open_file_excel():
    """
    """
    #global df_return
    filename = filedialog.askopenfilename()
    try:
        df_test = pd.read_excel(filename)
        if all(elem in df_test.columns for elem in dataframe.LIST_COLUMNS):
            msg = "Le fichier "+str(filename)+" est prêt à être analysé"
            messagebox.showinfo(title="Import des données", message=msg)
        else:
            msg = "Le fichier "+str(filename)+" ne contient pas les paramètres demandés"
            messagebox.showinfo(title="Import des données", message=msg)
    except:
        msg = "Le fichier "+str(filename)+" n'est pas un fichier excel"
        messagebox.showinfo(title="Import des données", message=msg)
    else:
        source_excel_entry.insert(0, filename)
        #df = dataframe.dataframe(filename)
        #df.convert()
        #df_return = df.get_info_user_relance()
    #print(df_return)
     
# Verify trees
def verify_trees_button():
    """
    Retourne un message d'information sur le nombre d'arbres pour \
    lesquels la relance est à faire
    """
    #Excel = source_excel_entry.get()
    # Messagebox
    if combo.get() != '':
        date_debut, date_fin = get_date()
        
        try:
            print(date_debut)
            print(date_fin)
            filename = source_excel_entry.get()
            print(filename)
            df = dataframe.dataframe(filename, date_debut, date_fin)
            df.convert()
            print("test convert")
            global df_return 
            df_return = df.get_info_user_relance()
            print("test return")
            print(len(df_return))
            #msg = "Le fichier excel a bien été analysé"
            #messagebox.showinfo(title="Information", message=msg)
            
        except:
            msg = "Il n'y a pas de relance à effectuer"
            messagebox.showinfo(title="Information", message=msg)

        else:
            if len(df_return) > 1:
                msg = "Il y a "+str(len(df_return))+" arbres pour lesquels une relance est à effectuer"
            elif len(df_return) == 1:
                msg = "Il y a un arbre pour lequel une relance est à effectuer"
            #else:
            #    msg = "Il n'y a pas de relance à effectuer"
            #messagebox.showinfo(title="Information", message=msg)
    else:
        msg = "Veuillez choisir la date de début et la date de fin pour la période de vérification des relances à considérer."
        messagebox.showerror(title="Erreur", message=msg)
        
            
    
def produce_mail_button():
    """
    Retourne un message d'informations sur le nombre de clients \
    pour qui on va créer un brouillon
    """
    if email_entry.get() != '' and password_entry.get() != '':
        # Creer objet mail
        #mail_return = mail.mail(email_entry.get(), password_entry.get(), df_return)
        mail_return = mail.mail(df_return)
        msg = "L'email et le mot de passe ont bien été renseignés"
        messagebox.showinfo(title="Information ", message=msg)          
            
        try:
            if len(mail_return.dataframe_user) != 0:
                if len(mail_return.dataframe_user) > 1:
                    msg = "Vous allez créer "+str(len(mail_return.dataframe_user))+" brouillons"
                    messagebox.askokcancel(title="Génération de mail", message=msg)
                    mail_return.create_draft()
                    # étape de vérification
                    msg = "Les brouillons ont bien été créés"
                    messagebox.showinfo(title="Génération de mail", message=msg)
    
                elif len(mail_return.dataframe_user) == 1:
                    msg = "Vous allez créer un brouillon"
                    messagebox.askokcancel(title="Génération de mail", message=msg)
                    mail_return.create_draft()
                    # étape de vérification
                    msg = "Le brouillon a bien été créé"
                    messagebox.showinfo(title="Génération de mail", message=msg)   
    
                # save last utilisation date in db
                file = open("db.txt","a") 
                l = max(df_return['Date de relance']).strftime('%m/%d/%Y')
                file.write(l+"\n") 
                file.close()
                    
            else:
                msg = "Il n'y a aucun mail à générer"
                messagebox.showinfo(title="Génération de mail", message=msg)
                
        except:
            msg = "Il faut appuyer sur le bouton 'Vérification des arbres' d'abord."
            messagebox.showerror(title="Erreur", message=msg)
    else:
        msg = "Merci de renseigner les champs 'email' et 'mot de passe' svp"
        messagebox.showerror(title="Erreur", message=msg)
                
def send_mail_button():
    """
    Retourne un message d'informations sur le nombre de clients \
    pour qui on va envoyer un mail
    """
    if email_entry.get() != '' and password_entry.get() != '':
        # Creer objet mail
        #mail_return = mail.mail(email_entry.get(), password_entry.get(), df_return)
        mail_return = mail.mail(df_return)
        msg = "L'email et le mot de passe ont bien été renseignés"
        messagebox.showinfo(title="Information ", message=msg)
        try:
            if len(mail_return.dataframe_user) != 0:
                if len(mail_return.dataframe_user) > 1:
                    msg ="Vous allez envoyer "+str(len(mail_return.dataframe_user))+" mails"
                    messagebox.askokcancel(title="Envoi de mail", message=msg)
                    mail_return.send_emails()
                    # étape de vérification
                    msg = "Les mails ont bien été envoyés"
                    messagebox.showinfo(title="Envoi de mail", message=msg)
                elif len(mail_return.dataframe_user) == 1:
                    msg ="Vous allez envoyer un mail"
                    messagebox.askokcancel(title="Envoi de mail", message=msg)
                    mail_return.send_emails()
                    # étape de vérification
                    msg = "Le mail a bien été envoyé"
                    messagebox.showinfo(title="Envoi de mail", message=msg)
    
                # save last utilisation date in db
                file = open("db.txt","a") 
                l = max(df_return['Date de relance']).strftime('%m/%d/%Y')
                file.write(l+"\n") 
                file.close()
                
            else:
                msg = "Il n'y a aucun mail à envoyer"
                messagebox.showinfo(title="Envoi de mail", message=msg)
        
        except:
            msg = "Il faut appuyer sur le bouton 'Vérification des arbres' d'abord."
            messagebox.showerror(title="Erreur", message=msg)
    else:
        msg = "Merci de renseigner les champs 'email' et 'mot de passe' svp"
        messagebox.showerror(title="Erreur", message=msg)
    
# Create window
window = Tk()
window.title("Arbre Conseil ® - Relances clients")
window.config(padx=24, pady=24)

# Canvas
canvas = Canvas(master=window, width=200, height=200)
logo_img = PhotoImage(file="logo.png")
canvas.create_image(100,100,anchor=CENTER,image=logo_img)
canvas.grid(row=0,column=0,columnspan=2)

# Labels
source_excel = Label(text="Fichier Excel :  ", justify=LEFT, anchor="w")
source_excel.grid(row=1, column=0, sticky=W)
email_label = Label(text="Email outlook :  ", justify=LEFT, anchor="w")
email_label.grid(row=7, column=0, sticky=W)
password_label = Label(text="Mot de passe :  ", justify=LEFT, anchor="w")
password_label.grid(row=8, column=0, sticky=W)

lastmail = Label(text="Dernière relance :  ", justify=LEFT, anchor="w")
lastmail.grid(row=3, column=0, sticky=W)
fileHandle = open ( 'db.txt',"r" )
lastmail_value_text = fileHandle.readlines()
fileHandle.close()
lastmail_value_text = (datetime.strptime(lastmail_value_text[len(lastmail_value_text)-1][:-1], '%m/%d/%Y')).strftime('%m/%d/%Y')
lastmail_value = Label(text=str(lastmail_value_text))
lastmail_value.grid(row=3, column=1)

combo = Label(text="Vérifier relances depuis :  ", justify=LEFT, anchor="w")
combo.grid(row=4, column=0, sticky=W)
until = Label(text="Jusqu'à :  ", justify=LEFT, anchor="w")
until.grid(row=5, column=0, sticky=W)
until_value = Label(text="Aujourd'hui : (" + datetime.today().date().strftime('%m/%d/%Y') + ")")
until_value.grid(row=5, column=1)

# Entries
source_excel_entry = Entry(width=35)
source_excel_entry.grid(row=1, column=1)
email_entry = Entry(width=35)
email_entry.grid(row=7,column=1)
email_entry.focus()
password_entry = Entry(width=35, show='*')
password_entry.grid(row=8, column=1)

# Combobox
combo = ttk.Combobox(values=["Dernière relance", "1 semaine","2 semaines","3 semaines","1 mois","2 mois","3 mois"], width=32, state="readonly")
combo.grid(row=4, column=1)

# Button
# Open file excel
open_file_font = font.Font(size=8, weight='bold')
open_file_button = Button(text="Ouvrir fichier Excel",bd=3,font=open_file_font,width=29,command=open_file_excel)
open_file_button.grid(row=2, column=1)
# Verify contenu file Excel
verify_button = Button(text="Vérification des arbres",bd=3,width=29,command=verify_trees_button)
verify_button.grid(row=6, column=1)
# Create crafts
produce_mail_button = Button(text="Créer des brouillons de mails",bd=3,width=29,command=produce_mail_button)
produce_mail_button.grid(row=9, column=1)
# Send mail
send_mail_button = Button(text="Créer et envoyer des mails",bd=3,width=29,command=send_mail_button)
send_mail_button.grid(row=10, column=1)

window.mainloop()