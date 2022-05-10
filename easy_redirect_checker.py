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
    plan.columns.values[0] = "Source URL"
    plan.columns.values[1] = "Target URL"
    plan = plan[['Source URL', 'Target URL']]
    onglet1 = pd.merge(plan, redirection[["Adresse", "Code de statut 1", 'Nombre de redirections', 'Adresse finale', 'Code de statut final', 'Boucle']], how='outer', left_on='Source URL', right_on='Adresse')
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
    onglet1[['Source URL', 'Target URL', 'Adresse finale', 'Code de statut 1', 'Boucle', 'Nombre de redirections', 'Code de statut final', 'Match URL', 'Validation', 'Remarque']]
    onglet2 = onglet1[['Validation','Remarque']]
    onglet2 = onglet2.value_counts().reset_index()
    onglet2 = onglet2.rename(columns={0: "Nombre URLs"})
    df = onglet1[['Remarque']]
    df = df.value_counts(normalize=True).reset_index()
    df = df.rename(columns={0: "Pourcentage URL"})
    onglet2 = pd.merge(onglet2, df, how='outer', left_on='Remarque', right_on='Remarque')
    chemin_dossier = os.path.dirname(chemin_fichier)
    if "test_de_redirections.xlsx" in os.listdir(chemin_dossier):
        else :
        with pd.ExcelWriter(str(chemin_dossier) + "\\" + "test_de_redirections.xlsx", engine='xlsxwriter') as writer:
            plan.to_excel(writer, sheet_name='Redirect Plan', index=False)
            redirection.to_excel(writer, sheet_name='Export SF', index=False)
            onglet1.to_excel(writer, sheet_name='Current Redirect', index=False)
            onglet2.to_excel(writer, sheet_name='Synthese')
            workbook  = writer.book
            worksheet = writer.sheets['Synthese']
            format2 = workbook.add_format({'num_format': '0%'})
            worksheet.set_column(4, 4, None, format2)
   
        
