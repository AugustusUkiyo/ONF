 #!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Dec  9 12:45:41 2020

@author: ukiyo
"""

import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta

LIST_COLUMNS = ['GlobalID', 'Nom du client', 'Coordonnées mail', 'Site', 'Numéro de l\'arbre', 
                   "Date du relevé", "Nom français de l\'arbre", 'Type de contrôle/suivi', 
                   'Date de relance pour contrôle/suivi', "Type d'intervention 1", 
                   'Date de relance pour intervention 1', "Type d'intervention 2", 
                   'Date de relance pour intervention 2', 'x', 'y']
#file_excel = 'Excel_type.xlsx'

class dataframe:
    def __init__(self, file_excel, date_debut, date_fin):
        self.df = pd.read_excel(file_excel) # sheet_name='DVS_v2_0'
        self.df_final = self.df[LIST_COLUMNS].copy() 
        # ordonner le dataframe par nom de client puis par date de relevé et enfin par numéro d'arbre
        self.date_debut = datetime.strptime(date_debut, '%m/%d/%Y').date()
        self.date_fin = datetime.strptime(date_fin, '%m/%d/%Y').date()
        self.date_today = datetime.today().date()
        self.date_today_str = datetime.today().strftime('%m/%d/%Y')

        print("test init")
        print(self.df_final)
    
    # Convert
    def convert(self):

        date1 = [ x.date() for x in list(self.df_final['Date de relance pour contrôle/suivi'])]
        date2 = [ x.date() for x in list(self.df_final['Date de relance pour intervention 1'])]
        date3 = [ x.date() for x in list(self.df_final['Date de relance pour intervention 2'])]
 
        # Verifier les dates de relances
        bool_date1 = pd.Series(np.array([ self.date_debut <= y and self.date_fin >= y for y in date1]))
        bool_date2 = pd.Series(np.array([ self.date_debut <= y and self.date_fin >= y for y in date2]))
        bool_date3 = pd.Series(np.array([ self.date_debut <= y and self.date_fin >= y for y in date3]))

        date_relance = []
        type_operation = []
        for i in range(len(bool_date1)):
            if bool_date1[i] == True and bool_date2[i] == True and bool_date3[i] == True:
                date_relance.append(min(date1[i],date2[i],date3[i]))
                type_operation.append(self.df_final.iloc[i]["Type de contrôle/suivi"]+" / "+self.df_final.iloc[i]["Type d'intervention 1"]+" / "+self.df_final.iloc[i]["Type d'intervention 2"])
            elif bool_date1[i] == True and bool_date2[i] == True:
                date_relance.append(min(date1[i],date2[i]))
                type_operation.append(self.df_final.iloc[i]["Type de contrôle/suivi"]+" / "+self.df_final.iloc[i]["Type d'intervention 1"])
            elif bool_date1[i] == True and bool_date3[i] == True:
                date_relance.append(min(date1[i],date3[i]))
                type_operation.append(self.df_final.iloc[i]["Type de contrôle/suivi"]+" / "+self.df_final.iloc[i]["Type d'intervention 2"])
            elif bool_date2[i] == True and bool_date3[i] == True:
                date_relance.append(min(date2[i],date3[i]))
                type_operation.append(self.df_final.iloc[i]["Type d'intervention 1"]+" / "+self.df_final.iloc[i]["Type d'intervention 2"])
            elif bool_date1[i] == True :
                date_relance.append(date1[i])
                type_operation.append(self.df_final.iloc[i]["Type de contrôle/suivi"])
            elif bool_date2[i] == True :
                date_relance.append(date2[i])
                type_operation.append(self.df_final.iloc[i]["Type d'intervention 1"])
            elif bool_date3[i] == True :
                date_relance.append(date3[i])
                type_operation.append(self.df_final.iloc[i]["Type d'intervention 2"])
            else:
                date_relance.append(None)
                type_operation.append('')
        
        date_relance = pd.Series(date_relance)
        type_operation = pd.Series(type_operation)
        print(date_relance, type_operation)

        # Ajouter dans la colonne de dataframe
        bool_relance = bool_date1 | bool_date2 | bool_date3

        self.df_final['Relance'] = list(bool_relance)
        print("relance ok: ", self.df_final['Relance'])
        self.df_final['Date de relance'] = date_relance
        print("deadline ok: ", self.df_final['Date de relance'])
        self.df_final["Type d'opération"] = type_operation
        print("operation ok")
        self.df_final['Deadline'] = (self.df_final['Date de relance'] + timedelta(days=90))
        print("deadline ok")

        print(self.df_final)

        #print(self.df_final['Relance mail ce jour '+self.date_today_str])
    
    def get_info_user_relance(self):
        # Retourner tous les infos du clients aui ont le date de relance aujourd'hui
        print(self.df_final[self.df_final['Relance']==True])
        return self.df_final[self.df_final['Relance']==True]
