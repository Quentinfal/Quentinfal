#!/usr/bin/env python
# coding: utf-8

# In[22]:


#!/usr/bin/env python
# coding: utf-8

# In[16]:
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import tkinter.font as font
import pandas as pd
import numpy as np
import xlsxwriter
import os
from pandas.io.json import json_normalize
from pandas.io import gbq
import re
import glob
from datetime import datetime
import json
from tkinter.ttk import Progressbar
from tkinter import *
from tkinter.ttk import *
def popupmsg(msg):
    popup = tk.Tk()
    popup.wm_title("Votre attention s'il vous plait !")
    popup.configure(background="#FA7F7F")
    label_popup = tk.Label(popup, text=msg, foreground = "#641E16", background = "#FA7F7F")
    label_popup['font'] = f_label
    label_popup.pack()
    bouton_popup = tk.Button(popup, text="C'est compris!", command = popup.destroy)
    bouton_popup.configure(foreground="#641E16", background = "white")
    bouton_popup['font'] = f_bouton
    bouton_popup.pack()
    popup.mainloop()
    
def traitement():
    chemin_fichier = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xlsx')), ('CSV', '*.csv')])
    if chemin_fichier.endswith('.csv'):
        plan = pd.read_csv(chemin_fichier)
    else:
        plan = pd.read_excel(chemin_fichier, engine='openpyxl')
    chemin_fichier2 = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xlsx')), ('CSV', '*.csv')])
    if chemin_fichier2.endswith('.csv'):
        redirection = pd.read_csv(chemin_fichier2)
    else:
        redirection = pd.read_excel(chemin_fichier2, engine='openpyxl')
        
     ####################    VOTRE CODE    #####################
    plan.columns.values[0] = "Source URL"
    plan.columns.values[1] = "Target URL"
    Redirection_pour_onglet1 = redirection[["Adresse", "Code de statut 1", 'Nombre de redirections', 'Adresse finale', 'Code de statut final', 'Boucle']]
    onglet1 = pd.merge(plan, Redirection_pour_onglet1, how='outer', left_on='Source URL', right_on='Adresse')
    onglet1['Match URL'] = np.where(onglet1['Target URL']  == onglet1['Adresse finale'], True, False)
    onglet1['Validation'] = np.where((onglet1['Match URL']  == True) & (onglet1['Code de statut final'] == 200), 'Ok Perfect', "3 - Problem detected")
    onglet1['Validation'] = np.where((onglet1['Match URL']  == True) & (onglet1['Code de statut final'] == 200) & (onglet1['Nombre de redirections'] == 2), 'Ok', onglet1['Validation'])
    onglet1['Remarque'] = np.where(onglet1['Validation']  == 'Ok Perfect', 'RAS', "Autre")
    onglet1['Remarque'] = np.where(onglet1['Validation']  == 'Ok', '2 redirections', onglet1['Remarque'])
    onglet1['Remarque'] = np.where((onglet1['Code de statut 1']  == 404) & (onglet1['Nombre de redirections'] == 0), 'Non redirigée', onglet1['Remarque'])
    onglet1['Remarque'] = np.where((onglet1['Code de statut final']  == 200) & (onglet1['Match URL'] == False), 'Ne redirige pas vers URL du plan', onglet1['Remarque'])
    onglet1['Remarque'] = np.where(onglet1['Nombre de redirections']  > 2, 'Chaine de redirection', onglet1['Remarque'])
    onglet1['Remarque'] = np.where(onglet1['Boucle']  > True, 'Boucle de redirection', onglet1['Remarque'])
    onglet1['Remarque'] = np.where((onglet1['Code de statut final']  == 404) & (onglet1['Nombre de redirections']  != 0), 'Redirige vers 404', onglet1['Remarque'])
    onglet1['Remarque'] = np.where((onglet1['Code de statut final']  == 500) & (onglet1['Nombre de redirections']  != 0), 'Redirige vers 500', onglet1['Remarque'])
    df = onglet1[['Validation','Remarque']]
    df = df.value_counts().reset_index()
    df = df.rename(columns={0: "Nombre URLs"})
    df2 = onglet1[['Remarque']]
    df2 = df2.value_counts(normalize=True).reset_index()
    df2 = df2.rename(columns={0: "Pourcentage URL"})
    onglet2 = pd.merge(df, df2, how='outer', left_on='Remarque', right_on='Remarque')
       
    ##################  FIN #########################
    chemin_dossier = os.path.dirname(chemin_fichier)
    if "test_de_redirections.xlsx" in os.listdir(chemin_dossier):
        popupmsg("Il y a déjà un fichier nommé 'test_de_redirections.xlsx' dans le dossier," + "\n"  + "\n" + "veuillez renommer ce fichier pour que le programme puisse en générer un nouveau.")
    else :
        with pd.ExcelWriter(str(chemin_dossier) + "\\" + "test_de_redirections.xlsx", engine='xlsxwriter') as writer:
            plan.to_excel(writer, sheet_name='Redirect Plan', index=False)
            redirection.to_excel(writer, sheet_name='Export SF', index=False)
            onglet1.to_excel(writer, sheet_name='Current Redirect', index=False)
            onglet2.to_excel(writer, sheet_name='Synthese', index=False)
        popupmsg("Votre fichier excel est généré =) Vous pouvez le trouver dans le même dossier que le fichier excel initial." + "\n" + "\n" + "A bientôt!" + "\n")
        
interface = tk.Tk()
interface.iconbitmap('icon4.ico')
interface.title("Easy Redirect Checker - For Screaming Frog")
interface.configure(background="#f2fafc")
# 1er message
f_label = font.Font(family='Arial', size=10)
f_bouton = font.Font(family='Arial', size=12, weight="bold")
label = tk.Label(text="\n"  + "Bienvenue dans Easy Redirect Checker - For Screaming Frog !" + "\n"  + "\n" + "Cet outil permet de recetter votre plan de redirections" + "\n"  + "\n" + "Veuillez insérer votre plan en Excel ou CSV et votre export chaines de redirection Screaming Frog :" + "\n" , foreground = "#3E4446", background = "#f2fafc")
label['font'] = f_label
label.pack(expand="yes")
# 1er bouton
bouton = tk.Button(text='Importer le plan de redirection puis votre Export SF ', command = traitement)
bouton.place(relx=0.200, rely=0.06, height=500, width=147)
bouton.configure(foreground="#3E4446")
bouton['font'] = f_bouton
bouton.pack(expand="yes")
# 2eme message
label2 = tk.Label(text="\n"  + "Une fois les fichiers chargés, un nouveau fichier Excel (nommé 'test_de_redirections.xlsx')"  + "\n" + "sera créé dans le même dossier que le premier fichier séléctionné." + "\n", foreground = "#3E4446", background = "#f2fafc")
label2['font'] = f_label
label2.pack(expand='yes')
# 4eme bouton
bouton4 = tk.Button(interface, text="Fermer", command = interface.destroy)
bouton4.configure(foreground="#3E4446")
bouton4['font'] = f_bouton
bouton4.pack(expand="yes")
interface.mainloop()

