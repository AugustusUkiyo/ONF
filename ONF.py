#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Dec  7 15:01:20 2020

@author: ukiyo
"""



import pandas as pd
import numpy as np
from datetime import datetime


#df = pd.read_csv('fichier_Excel.csv', encoding = 'utf8')
df = pd.read_excel('Excel_type.xlsx', sheet_name='DVS_v2_0')
df_final = df[['GlobalID', 'Nom du client', 'Coordonnées mail', 'Numéro de l\'arbre', "Date du relevé", "Nom français de l\'arbre", 'Type de contrôle/suivi', 'Date de relance pour contrôle/suivi', "Type d'intervention 1", 'Date de relance pour intervention 1', "Type d'intervention 2", 'Date de relance pour intervention 2', 'x', 'y']]
date_today = datetime.now().date()

mask_list = [   df_final['Date de relance pour contrôle/suivi'] == date_today, 
                df_final['Date de relance pour intervention 1'] == date_today, 
                df_final['Date de relance pour intervention 2'] == date_today
            ]

combined_mask = np.array(sum(mask_list), dtype=np.bool)

matching_date_rows = df_final[combined_mask]

print(matching_date_rows)