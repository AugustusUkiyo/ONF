#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Dec  9 12:45:41 2020

@author: ukiyo
"""

import pandas as pd
import numpy as np
from datetime import datetime

LIST_COLUMNS = ['GlobalID', 'Nom du client', 'Coordonnées mail', 'Numéro de l\'arbre', 
                   "Date du relevé", "Nom français de l\'arbre", 'Type de contrôle/suivi', 
                   'Date de relance pour contrôle/suivi', "Type d'intervention 1", 
                   'Date de relance pour intervention 1', "Type d'intervention 2", 
                   'Date de relance pour intervention 2', 'x', 'y']
#file_excel = 'Excel_type.xlsx'

class dataframe:
    def __init__(self, file_excel,date_debut, date_fin):
        self.df = pd.read_excel(file_excel) # sheet_name='DVS_v2_0'
        self.df_final = self.df[LIST_COLUMNS]
        self.date_debut = datetime.strptime(date_debut, '%m/%d/%Y').date()
        self.date_fin = datetime.strptime(date_fin, '%m/%d/%Y').date()
        self.date_today = datetime.today().date()
        self.date_today_str = datetime.today().strftime('%m/%d/%Y')
    
    # Convert
    def convert(self):
        
        date1 = [ x.date() for x in list(self.df_final['Date de relance pour contrôle/suivi'])]
        date2 = [ x.date() for x in list(self.df_final['Date de relance pour intervention 1'])]
        date3 = [ x.date() for x in list(self.df_final['Date de relance pour intervention 2'])]
        
        # Verifier les dates de relances
        bool_date1 = np.array([ self.date_debut <= y and self.date_fin >= y for y in date1])
        bool_date2 = np.array([ self.date_debut <= y and self.date_fin >= y for y in date2])
        bool_date3 = np.array([ self.date_debut <= y and self.date_fin >= y for y in date3])
        
        # Ajouter dans la colonne de dataframe
        bool_relance = bool_date1 + bool_date2 + bool_date3
        self.df_final['Relance mail ce jour '+self.date_today_str] = bool_relance
        #print(self.df_final['Relance mail ce jour '+self.date_today_str])
    
    def get_info_user_relance(self):
        # Retourner tous les infos du clients aui ont le date de relance aujourd'hui
        #print(self.df_final[self.df_final['Relance mail ce jour '+self.date_today_str]==True])
        return self.df_final[self.df_final['Relance mail ce jour '+self.date_today_str]==True]
