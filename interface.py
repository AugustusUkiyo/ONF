#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Dec  9 12:45:40 2020

@author: ukiyo
"""


from tkinter import *
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
    if combo1.get() == 'Date actuelle':
        date_debut = datetime.today().strftime('%m/%d/%Y')
    else:
        # read db to get last utilisation
        fileHandle = open ( 'db.txt',"r" )
        lineList = fileHandle.readlines()
        fileHandle.close()
        date_debut = (datetime.strptime(lineList[len(lineList)-1][:-1], '%m/%d/%Y')).strftime('%m/%d/%Y')
    
    if combo2.get() =='Date actuelle':
        date_fin = datetime.today().date().strftime('%m/%d/%Y')
    else:
        
        if combo2.get() =="1 semaine":
            date_fin = (datetime.today().date() + timedelta(days=7)).strftime('%m/%d/%Y')
        if combo2.get() =="2 semaines":
            date_fin = (datetime.today().date() + timedelta(days=14)).strftime('%m/%d/%Y')
        if combo2.get() =="3 semaines":
            date_fin = (datetime.today().date() + timedelta(days=21)).strftime('%m/%d/%Y')
        if combo2.get() =="1 mois":
            date_fin = (datetime.today().date() + timedelta(days=30)).strftime('%m/%d/%Y')
        if combo2.get() =="2 mois":
            date_fin = (datetime.today().date() + timedelta(days=60)).strftime('%m/%d/%Y')
        if combo2.get() =="3 mois":
            date_fin = (datetime.today().date() + timedelta(days=120)).strftime('%m/%d/%Y')
    
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
    if combo1.get() != '' and combo2.get() != '':
        date_debut, date_fin = get_date()
        
        try:
            #print(date_debut)
            #print(date_fin)
            filename = source_excel_entry.get()
            #print("ok 1")
            df = dataframe.dataframe(filename, date_debut, date_fin)
            df.convert()
            global mail_return 
            df_return = df.get_info_user_relance()
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
            
            if email_entry.get() != '' and password_entry.get() != '':
                # Creer objet mail
                #mail_return = mail.mail(email_entry.get(), password_entry.get(), df_return)
                mail_return = mail.mail(df_return)
                msg = "L'email et le mot de passe ont bien été renseignés"
                messagebox.showinfo(title="Information ", message=msg)          
            else:
                msg = "Merci de renseigner les champs 'email' et 'mot de passe' svp"
                messagebox.showerror(title="Erreur", message=msg)

    else:
        msg = "Veuillez choisir la date de début et la date de fin pour la période de vérification des relances à considérer."
        messagebox.showerror(title="Erreur", message=msg)
        
            
    
def produce_mail_button():
    """
    Retourne un message d'informations sur le nombre de clients \
    pour qui on va créer un brouillon
    """
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
            l = datetime.today().strftime('%m/%d/%Y')
            file.write(l+"\n") 
            file.close()
                
        else:
            msg = "Il n'y a aucun mail à générer"
            messagebox.showinfo(title="Génération de mail", message=msg)
            
    except:
        msg = "Il faut appuyer sur le bouton 'Vérification des arbres' d'abord."
        messagebox.showerror(title="Erreur", message=msg)
    
def send_mail_button():
    """
    Retourne un message d'informations sur le nombre de clients \
    pour qui on va envoyer un mail
    """
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
            l = datetime.today().strftime('%m/%d/%Y')
            file.write(l+"\n") 
            file.close()
            
        else:
            msg = "Il n'y a aucun mail à envoyer"
            messagebox.showinfo(title="Envoi de mail", message=msg)
    
    except:
        msg = "Il faut appuyer sur le bouton 'Vérification des arbres' d'abord."
        messagebox.showerror(title="Erreur", message=msg)
    
# Create window
window = Tk()
window.title("Arbre Conseil ® - Relances clients")
window.config(padx=20, pady=20)

# Canvas
canvas = Canvas(width=200, height=200)
logo_img = PhotoImage(file="logo.png")
canvas.create_image(100,100,image=logo_img)
canvas.grid(row=0,column=1)

# Labels
source_excel = Label(text="Fichier Excel")
source_excel.grid(row=1, column=0)
email_label = Label(text="Email outlook")
email_label.grid(row=2, column=0)
password_label = Label(text="Mot de passe")
password_label.grid(row=3, column=0)
combo1 = Label(text="Date de début")
combo1.grid(row=4, column=0)
combo2 = Label(text="Date de fin")
combo2.grid(row=5, column=0)


# Entries
source_excel_entry = Entry(width=35)
source_excel_entry.grid(row=1, column=1)
email_entry = Entry(width=35)
email_entry.grid(row=2,column=1)
email_entry.focus()
password_entry = Entry(width=35, show='*')
password_entry.grid(row=3, column=1)

# Combobox
combo1 = ttk.Combobox(values=["Dernière utilisation", "Date actuelle"],width=34)
combo1.grid(row=4, column=1)
combo2 = ttk.Combobox(values=["Date actuelle","1 semaine","2 semaines","3 semaines","1 mois","2 mois","3 mois"],width=34)
combo2.grid(row=5, column=1)

# Button
# Open file excel
open_file_button = Button(text="Ouvrir fichier Excel",command=open_file_excel)
open_file_button.grid(row=1, column=3)
# Verify contenu file Excel
verify_button = Button(text="Vérification des arbres",width=25,command=verify_trees_button)
verify_button.grid(row=6, column=1)
# Create crafts
produce_mail_button = Button(text="Créer des brouillons de mails",width=25,command=produce_mail_button)
produce_mail_button.grid(row=7, column=1)
# Send mail
send_mail_button = Button(text="Créer et envoyer des mails",width=25,command=send_mail_button)
send_mail_button.grid(row=8, column=1)

window.mainloop()