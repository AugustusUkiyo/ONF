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
    if combo1.get() =='current date':
        date_debut = datetime.today().strftime('%m/%d/%Y')
    else:
        # read db to get last utilisation
        fileHandle = open ( 'db.txt',"r" )
        lineList = fileHandle.readlines()
        fileHandle.close()
        date_debut = (datetime.strptime(lineList[len(lineList)-1][:-1], '%m/%d/%Y')).strftime('%m/%d/%Y')
    
    if combo2.get() =='current date':
        date_fin = datetime.today().date()
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
            msg = "Le fichier "+str(filename)+" est pret a analyser"
            messagebox.showinfo(title="Information", message=msg)
        else:
            msg = "Le fichier "+str(filename)+" ne contient pas les parametres demandes"
            messagebox.showinfo(title="Information", message=msg)
    except:
        msg = "Le fichier "+str(filename)+" n'est pas un fichier excel"
        messagebox.showinfo(title="Information", message=msg)
    else:
        source_excel_entry.insert(0, filename)
        #df = dataframe.dataframe(filename)
        #df.convert()
        #df_return = df.get_info_user_relance()
    #print(df_return)
     
# Verify user
def verify_user_button():
    """
    Retourner le nombre de personne dont on va envoyer le mail
    """
    #Excel = source_excel_entry.get()
    # Messagebox
    if combo1.get() != '' and combo2.get() != '':
        date_debut,date_fin = get_date()
        
        try:
            print(date_debut)
            print(date_fin)
            filename = source_excel_entry.get()
            print("ok 1")
            df = dataframe.dataframe(filename, date_debut, date_fin)
            df.convert()
            global mail_return 
            df_return = df.get_info_user_relance()
            msg = "fichier excel est charge"
            messagebox.showinfo(title="Information", message=msg)
            
        except:
            msg = "ouvrir le fichier excel svp"
            messagebox.showerror(title="Information", message=msg)
        else:
            if len(df_return) > 1:
                msg = "Il y a "+str(len(df_return))+" personnes qui ont la date de relancement aujourd'hui"
            else:
                msg = "Il y a "+str(len(df_return))+" personne qui ont la date de relancement aujourd'hui"
            messagebox.showinfo(title="Information", message=msg)
            
            if email_entry.get() != '' and password_entry.get() != '':
                # Creer objet mail
                #mail_return = mail.mail(email_entry.get(), password_entry.get(), df_return)
                mail_return = mail.mail(df_return)
                msg = "email et le mot de passe ont bien rempli"
                messagebox.showinfo(title="Information ", message=msg)          
            else:
                msg = "remplir email et le mot de passe svp"
                messagebox.showerror(title="Information", message=msg)
    else:
        msg = "choisir la date de debut et date fin svp"
        messagebox.showerror(title="Information", message=msg)
        
            
    
def produce_mail_button():
    """
    Retourner le nombre de personne dont on va envoyer le mail
    """
    try:
        if len(mail_return.dataframe_user) > 0:
            msg ="Vous allez creer des brouillons"
            messagebox.askokcancel(title="??", message=msg)
            mail_return.create_draft()
    
            msg = "brouillons ont reussi a creer"
            messagebox.showinfo(title="Information", message=msg)
            # save last utilisation date in db
            file = open("db.txt","a") 
            l = datetime.today().strftime('%m/%d/%Y')
            file.write(l+"\n") 
            file.close()
                
            
        else:
            msg = "Il y a "+str(len(mail_return.dataframe_user))+" personne qui ont la date de relancement aujourd'hui"
            messagebox.showinfo(title="Information", message=msg)
    except:
        msg = "Appuyer sur le button verify user d'aborde"
        messagebox.showerror(title="Information", message=msg)
    
def sent_mail_button():
    """
    Retourner le nombre de personne dont on va envoyer le mail
    """
    try:
        if len(mail_return.dataframe_user) > 0:
            msg1 ="Vous allez envoyer des mails"
            messagebox.askokcancel(title="??", message=msg1)
            mail_return.send_emails()
            msg = "mails ont reussi a envoyer"
            messagebox.showinfo(title="Information", message=msg)
            # save last utilisation date in db
            file = open("db.txt","a") 
            l = datetime.today().strftime('%m/%d/%Y')
            file.write(l+"\n") 
            file.close()
            
        else:
            msg = "Il y a "+str(len(mail_return.dataframe_user))+" personne qui ont la date de relancement aujourd'hui"
            messagebox.showinfo(title="Information", message=msg)
    
    except:
        msg = "Appuyer sur le button verify user d'aborde"
        messagebox.showerror(title="Information", message=msg)
    
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
source_excel = Label(text="File Excel")
source_excel.grid(row=1, column=0)
email_label = Label(text="Email")
email_label.grid(row=2, column=0)
password_label = Label(text="Password")
password_label.grid(row=3, column=0)
combo1 = Label(text="Date debut")
combo1.grid(row=4, column=0)
combo2 = Label(text="Date fin")
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
combo1 = ttk.Combobox(values=["last utilisation","current date"],width=34)
combo1.grid(row=4, column=1)
combo2 = ttk.Combobox(values=["current date","1 semaine","2 semaines","3 semaines","1 mois","2 mois","3 mois"],width=34)
combo2.grid(row=5, column=1)

# Button
# Open file excel
open_file_button = Button(text="Open file Excel",command=open_file_excel)
open_file_button.grid(row=1, column=3)
# Verify contenu file Excel
verify_button = Button(text="Verify User",width=15,command=verify_user_button)
verify_button.grid(row=6, column=1)
# Create crafts
produce_mail_button = Button(text="Produce mail",width=15,command=produce_mail_button)
produce_mail_button.grid(row=7, column=1)
# Sent mail
sent_mail_button = Button(text="Sent mail",width=15,command=sent_mail_button)
sent_mail_button.grid(row=8, column=1)

window.mainloop()