#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Dec  9 12:45:40 2020

@author: ukiyo
"""


from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import dataframe
import pandas as pd 
#from tkFileDialog import askdirectory



#nb_pers = 2
#open file excel
def open_file_excel():
    """
    """
    global df_return
    filename = filedialog.askopenfilename()
    try:
        df_test = pd.read_excel(filename)
        if all(elem in df_test.columns for elem in dataframe.LIST_COLUMNS):
            msg = "Le fichier "+str(filename)+" est prêt à être analysé"
            messagebox.showinfo(title="Import des données", message=msg)
        else:
            msg = "Le fichier "+str(filename)+" ne contient pas les paramêtres demandés"
            messagebox.showinfo(title="Import des données", message=msg)
    except:
        msg = "Le fichier "+str(filename)+" n'est pas un fichier excel"
        messagebox.showinfo(title="Import des données", message=msg)
    else:
        source_excel_entry.insert(0, filename)
        df = dataframe.dataframe(filename)
        df.convert()
        df_return = df.get_info_user_relance()
    #print(df_return)
     
# Verify user
def verify_user_button():
    """
    Retourne un message d'information sur le nombre d'arbres pour lesquels la relance est à faire
    """
    #Excel = source_excel_entry.get()
    # Messagebox
    if len(df_return) > 1:
        msg = "Il y a "+str(len(df_return))+" arbres pour lesquels une relance est à effectuer"
    elif len(df_return) == 1:
        msg = "Il y a un arbre pour lequel une relance est à effectuer"
    else:
        msg = "Il n'y a pas de relance à effectuer"
    messagebox.showinfo(title="Information", message=msg)
    
def produce_mail_button():
    """
    Retourne un message d'informations sur le nombre de clients pour qui on va créer un brouillon
    """
    if len(df_return) > 1:
        msg = "Vous allez créer "+str(len(df_return))+" brouillons"
        messagebox.askokcancel(title="Génération de mail", message=msg)
        # étape de vérification
        msg = "Les brouillons ont bien été créés"
        messagebox.showinfo(title="Génération de mail", message=msg)
    elif len(df_return) == 1:
        msg = "Vous allez créer un brouillon"
        messagebox.askokcancel(title="Génération de mail", message=msg)
        # étape de vérification
        msg = "Le brouillon a bien été créé"
        messagebox.showinfo(title="Génération de mail", message=msg)
    else:
        msg = "Il n'y a aucun mail à générer"
        messagebox.showinfo(title="Génération de mail", message=msg)
    
def send_mail_button():
    """
    Retourne un message d'informations sur le nombre de clients pour qui on va envoyer un mail
    """
    if len(df_return) > 1:
        msg ="Vous allez envoyer "+str(len(df_return))+" mails"
        messagebox.askokcancel(title="Envoi de mail", message=msg)
        # étape de vérification
        msg = "Les mails ont bien été envoyés"
        messagebox.showinfo(title="Envoi de mail", message=msg)
    elif len(df_return) == 1:
        msg ="Vous allez envoyer un mail"
        messagebox.askokcancel(title="Envoi de mail", message=msg)
        # étape de vérification
        msg = "Le mail a bien été envoyé"
        messagebox.showinfo(title="Envoi de mail", message=msg)
    else:
        msg = "Il n'y a aucun mail à envoyer"
        messagebox.showinfo(title="Envoi de mail", message=msg)
    
# Create window
window = Tk()
window.title("ONF: Interface de relance client")
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

# Entries
source_excel_entry = Entry(width=35)
source_excel_entry.grid(row=1, column=1)
email_entry = Entry(width=35)
email_entry.grid(row=2,column=1)
password_entry = Entry(width=35, show='*')
password_entry.grid(row=3, column=1)

# Button
# Open file excel
open_file_button = Button(text="Ouvrir fichier Excel",command=open_file_excel)
open_file_button.grid(row=1, column=3)
# Verify contenu file Excel
verify_button = Button(text="Vérification des relances",width=25,command=verify_user_button)
verify_button.grid(row=4, column=1)
# Create crafts
produce_mail_button = Button(text="Créer des brouillons de mails",width=25,command=produce_mail_button)
produce_mail_button.grid(row=5, column=1)
# Sent mail
send_mail_button = Button(text="Créer et envoyer des mails",width=25,command=send_mail_button)
send_mail_button.grid(row=6, column=1)

window.mainloop()