#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Dec  7 15:01:20 2020

@author: ukiyo
"""



import pandas as pd
import numpy as np
import datetime as dt


#df = pd.read_csv('fichier_Excel.csv', encoding = 'utf8')
df = pd.read_excel('Excel_type.xlsx', sheet_name='DVS_v2_0')
df_final = df[['GlobalID', 'Nom du client', 'Coordonnées mail', 'Numéro de l\'arbre', "Date du relevé", "Nom français de l\'arbre", 'Type de contrôle/suivi', 'Date de relance pour contrôle/suivi', "Type d'intervention 1", 'Date de relance pour intervention 1', "Type d'intervention 2", 'Date de relance pour intervention 2', 'x', 'y']]

df_final['Date de relance pour contrôle/suivi'] = pd.to_datetime(df_final['Date de relance pour contrôle/suivi']).dt.date
df_final['Date de relance pour intervention 1'] = pd.to_datetime(df_final['Date de relance pour intervention 1']).dt.date
df_final['Date de relance pour intervention 2'] = pd.to_datetime(df_final['Date de relance pour intervention 2']).dt.date
# Ces trois lignes permettent de transformer les données au format timestamp en données au format datetime.date


date_today = dt.date.today()
test_date = dt.date(2020, 10, 15)

target = test_date
# permet de switcher entre un target du jour même et une valeur test pour vérifier que ça marche

mask_list = [   df_final['Date de relance pour contrôle/suivi'] == target, 
                df_final['Date de relance pour intervention 1'] == target, 
                df_final['Date de relance pour intervention 2'] == target
            ]
combined_mask = np.array(sum(mask_list), dtype=np.bool)
# j'ai groupé les masques mais maintenant il faut distinguer entre contrôle et intervention.

matching_date_rows = df_final[combined_mask]

print(df_final['Date de relance pour contrôle/suivi'])
print(df_final['Date de relance pour intervention 1'])
print(df_final['Date de relance pour intervention 2'])

print("today = ", date_today)
print("test =", test_date)

print(matching_date_rows)

# Il va falloir commencer à penser en terme de classes pour organiser le code selon les différents cas de figure