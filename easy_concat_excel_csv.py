#!/usr/bin/env python
# coding: utf-8

# In[13]:


#!/usr/bin/env python
# coding: utf-8

# In[16]:
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askopenfilenames
from tkinter.filedialog import askdirectory
import tkinter.font as font
import pandas as pd
import numpy as np
import os
import re
import glob

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
    
def popupmsg2(msg):
    popup = tk.Tk()
    popup.wm_title("Votre attention s'il vous plait !")
    popup.configure(background="#7ffa7f")
    label_popup = tk.Label(popup, text=msg, foreground = "#22780f", background = "#7ffa7f")
    label_popup['font'] = f_label
    label_popup.pack()
    bouton_popup = tk.Button(popup, text="C'est compris!", command = popup.destroy)
    bouton_popup.configure(foreground="#22780f", background = "white")
    bouton_popup['font'] = f_bouton
    bouton_popup.pack()
    popup.mainloop()
    
def traitement():
    chemin_fichier = askopenfilenames(filetypes=[('Excel', ('*.xls', '*.xlsx')), ('CSV', '*.csv')])
    df = pd.DataFrame()
    for chemin_fichier in chemin_fichier:
        data = pd.read_excel(chemin_fichier)
        df = df.append(data)
    chemin_dossier = os.path.dirname(chemin_fichier)
    if "fichier_concat.xlsx" in os.listdir(chemin_dossier):
        popupmsg("Il y a déjà un fichier nommé 'fichier_concat.xlsx' dans le dossier," + "\n"  + "\n" + "veuillez renommer ce fichier pour que le programme puisse en générer un nouveau.")
    else :
        with pd.ExcelWriter(str(chemin_dossier) + "\\" + "fichier_concat.xlsx", engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='feuille1', index=False)
        popupmsg2("Votre fichier excel est généré =) Vous pouvez le trouver dans le même dossier que le fichier excel initial." + "\n" + "\n" + "A bientôt!" + "\n")    

def traitement2():
    chemin_fichier = askopenfilenames(filetypes=[('CSV', '*.csv')])
    df = pd.DataFrame()
    for chemin_fichier in chemin_fichier:
        data = pd.read_csv(chemin_fichier)
        df = df.append(data)
    chemin_dossier = os.path.dirname(chemin_fichier)
    if "fichier_concat.csv" in os.listdir(chemin_dossier):
        popupmsg("Il y a déjà un fichier nommé 'fichier_concat.csv' dans le dossier," + "\n"  + "\n" + "veuillez renommer ce fichier pour que le programme puisse en générer un nouveau.")
    else :
        df.to_csv(str(chemin_dossier) + "\\" + "fichier_concat.csv", encoding="utf-8-sig", index=False)
        popupmsg2("Votre fichier excel est généré =) Vous pouvez le trouver dans le même dossier que le fichier excel initial." + "\n" + "\n" + "A bientôt!" + "\n")   
        
interface = tk.Tk()
interface.iconbitmap('icon.ico')
interface.title("EasyConcat'")
interface.configure(background="#f2fafc")
# 1er message
f_label = font.Font(family='Arial', size=10)
f_bouton = font.Font(family='Arial', size=11, weight="bold")
label = tk.Label(text="\n"  + "Bienvenue dans Easy Concat !" + "\n"  + "\n"  + "Concaténer vos Fichier CSV :" + "\n" , foreground = "#3E4446", background = "#f2fafc")
label['font'] = f_label
label.pack(expand="yes")
# 1er bouton
bouton = tk.Button(text='Importer les fichiers Excel', command = traitement)
bouton.place(relx=0.200, rely=0.06, height=500, width=147)
bouton.configure(foreground="#3E4446")
bouton['font'] = f_bouton
bouton.pack(expand="yes")
# 1er message
f_label = font.Font(family='Arial', size=10)
f_bouton = font.Font(family='Arial', size=11, weight="bold")
label = tk.Label(text="\n"  + "Concaténer vos Fichier CSV :" + "\n" , foreground = "#3E4446", background = "#f2fafc")
label['font'] = f_label
label.pack(expand="yes")
# 2er bouton
bouton2 = tk.Button(text='Importer les fichiers CSV', command = traitement2)
bouton2.place(relx=0.200, rely=0.06, height=500, width=147)
bouton2.configure(foreground="#3E4446")
bouton2['font'] = f_bouton
bouton2.pack(expand="yes")
# 2eme message
label2 = tk.Label(text="\n"  + "Un nouveau fichier (nommé 'fichier_concat.xlsx/csv')"  + "\n" + "sera créé dans votre dossier." + "\n", foreground = "#3E4446", background = "#f2fafc")
label2['font'] = f_label
label2.pack(expand='yes')
# 4eme bouton
bouton4 = tk.Button(interface, text="Fermer", command = interface.destroy)
bouton4.configure(foreground="#3E4446")
bouton4['font'] = f_bouton
bouton4.pack(expand="yes")
interface.mainloop()

