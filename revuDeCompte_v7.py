# -*- coding: utf-8 -*-
"""
Created on Fri May 13 11:17:37 2022

@author: 4180849
"""
import pandas as pd
import os
import csv
from datetime import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter.simpledialog import askstring
from tkinter import Tk
Tk().withdraw() 

pd.options.mode.chained_assignment = None  # default='warn'

# Script qui sort les chiffres pour la revue des comptes et les détails

# Rappels sur le fonctionnements des comptes SIDEP

# Un compte est mis en sommeil au bout de 90 jours d’inactivité
# Un compte ne peut pas être supprimé - il a un état : ACTIF (True /False)
# Lorsque on crée un compte, la date de dernier accès n’est pas renseignée : 
#    on distinguera les comptes « ACTIF » des comptes « Jamais accédé ». 


folder = askdirectory(title = 'Sélectionner dossier cible')
filename = askopenfilename(title='Fichier des données à traiter')
date_ext = askstring(title = 'Date extraction : AAAA/MM/DD',prompt='date_extraction : AAAA/MM/DD') 
date_ext = date_ext.split('/')
filename2 = askopenfilename(title='Fichier liste complète formulaire')
colnames = ["LoginName","Nom","Prenom","Telephone","Mail","Actif","Autre","DateDerniereConnexion",
            "Empty","Organisation","Groupe","NomGroupe","Cree_Par","Profile","DateCreation",'Autre2']
df = pd.read_csv(filename, sep=';', names = colnames, low_memory=False) 
df2 = pd.read_excel(filename2) 
os.chdir(folder)

# Nombre de compte sur SIDEP
nb_total = len(df['Actif'])-2      # on enlève 2 qui correspond aux 2 premières lignes             
nb_actif = len(df[df['Actif']=="T"])
nb_inactif = nb_total - nb_actif


df_actif = df[df['Actif']=="T"]
nb_jamais_accede = 0
nb_sommeil = 0
# Période de 90 jours à partir de la date de l'extract
troismois = datetime(int(date_ext[0]), int(date_ext[1]), int(date_ext[2])) - timedelta(days = 90)
liste_inactif = []
for i in df_actif.index:
    # cas des inactifs parmi les actifs
    if(df_actif['DateDerniereConnexion'][i] == '                   '):
        nb_jamais_accede += 1
        liste_inactif.append(i)
    # cas actif et en sommeil
    else:
        # formattage de la date
        date_test = df_actif['DateDerniereConnexion'][i]
        date_heure = date_test.split(' ')
        date = date_heure[0].split('/')
        heure = date_heure[1].split(':')
        dateDerniereCo = datetime(int(date[2]), int(date[1]), int(date[0]), int(heure[0]), int(heure[1]))
        if(dateDerniereCo <= troismois):
            nb_sommeil += 1
            liste_inactif.append(i)

df_vrai_actif = df_actif.copy()
df_vrai_actif.drop(labels = liste_inactif, axis = 0, inplace = True)

nb_inactif_total = nb_jamais_accede + nb_sommeil
nb_actif_total = nb_actif - nb_inactif_total

# Nombre de compte par organisation
compte_organisation = df_vrai_actif['Organisation']
c_orga = compte_organisation.value_counts().sort_index()
df_total_orga = pd.DataFrame([[c_orga.sum()]], index=['Total général'])
if not os.path.exists(folder + '/CompteParOrganisation'):
    os.makedirs(folder + '/CompteParOrganisation')
os.chdir(folder + '/CompteParOrganisation')
for i in c_orga.index:
    nom_orga = i.split(' ')
    nom_compte = "compte_" + str(nom_orga[0])
    df_nom_compte = df_vrai_actif[df_vrai_actif['Organisation']==i]
    df_nom_compte.to_csv(nom_compte + ".csv", sep=";", index=False)
os.chdir(folder)

# Détail des comptes PROSANTE
df_prosante = df_vrai_actif[df_vrai_actif['Organisation']=='PROSANTE                                                        ']
df_prosante['NomGroupe'] = df_prosante['NomGroupe'].str.upper()
df_prosante_rpps = df_prosante[df_prosante.NomGroupe.str.contains('rpps', case=False)]
df_infirmier = df_prosante[df_prosante.NomGroupe.str.contains('infirmier', case=False)]
df_biologiste = df_prosante[df_prosante.NomGroupe.str.contains('biologiste', case=False)]
df_pharmacien = df_prosante[df_prosante.NomGroupe.str.contains('pharmacien', case=False)]
df_medecin = df_prosante[df_prosante.NomGroupe.str.contains('medecin', case=False)]
df_dentiste = df_prosante[df_prosante.NomGroupe.str.contains('dentiste', case=False)]
df_masseur = df_prosante[df_prosante.NomGroupe.str.contains('kine', case=False)]
df_sage = df_prosante[df_prosante.NomGroupe.str.contains('sage', case=False)]

# Profession pour autres comptes créés manuellement
df_ambulancier = df_prosante[df_prosante.NomGroupe.str.contains('ambulancier', case=False)]
df_manipulateur = df_prosante[df_prosante.NomGroupe.str.contains('manipulateur', case=False)]
df_opticien = df_prosante[df_prosante.NomGroupe.str.contains('opticien', case=False)]
df_orthophoniste = df_prosante[df_prosante.NomGroupe.str.contains('orthophoniste', case=False)]
df_orthoptiste = df_prosante[df_prosante.NomGroupe.str.contains('orthoptiste', case=False)]
df_osteopathe = df_prosante[df_prosante.NomGroupe.str.contains('osteopathe', case=False)]
df_pedicure = df_prosante[df_prosante.NomGroupe.str.contains('pedicure', case=False)]
df_podologue = df_prosante[df_prosante.NomGroupe.str.contains('podologue', case=False)]
df_prothesiste = df_prosante[df_prosante.NomGroupe.str.contains('prothesiste', case=False)]
df_psychiatre = df_prosante[df_prosante.NomGroupe.str.contains('psychiatre', case=False)]

# Tableau des comptes crées manuellement par profession
df_manuellement_prosante = df_prosante[df_prosante['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_infirmier = df_infirmier[df_infirmier['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_biologiste = df_biologiste[df_biologiste['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_pharmacien = df_pharmacien[df_pharmacien['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_medecin = df_medecin[df_medecin['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_dentiste = df_dentiste[df_dentiste['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_masseur = df_masseur[df_masseur['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_sage = df_sage[df_sage['Cree_Par']!='                                                                                                                                                                                                                                                                ']

# Autres comptes PROSANTE créés manuellement
df_manuellement_ambulancier = df_ambulancier[df_ambulancier['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_manipulateur = df_manipulateur[df_manipulateur['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_opticien = df_opticien[df_opticien['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_orthophoniste = df_orthophoniste[df_orthophoniste['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_orthoptiste = df_orthoptiste[df_orthoptiste['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_osteopathe = df_osteopathe[df_osteopathe['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_pedicure = df_pedicure[df_pedicure['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_podologue = df_podologue[df_podologue['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_prothesiste = df_prothesiste[df_prothesiste['Cree_Par']!='                                                                                                                                                                                                                                                                ']
df_manuellement_psychiatre = df_psychiatre[df_psychiatre['Cree_Par']!='                                                                                                                                                                                                                                                                ']

# Autres comptes hors profession déjà traité
df_autres_total = df_prosante[df_prosante["NomGroupe"].str.contains('infirmier|biologiste|pharmacien|medecin|dentiste|kine|sage', case=False) == False]
df_manuellement_autres = df_manuellement_prosante[df_manuellement_prosante["NomGroupe"].str.contains('infirmier|biologiste|pharmacien|medecin|dentiste|kine|sage', case=False) == False]
df_manuellement_autres_X = df_manuellement_autres[df_manuellement_autres["NomGroupe"].str.contains('ambulancier|manipulateur|opticien|orthophoniste|orthoptiste|osteopathe|pedicure|podologue|prothesiste|psychiatre', case=False) == False]

# Autres profils que PROSANTE, colonne "Profile"
df_autre_profil_2prosante = df_prosante[df_prosante['Profile'].str.contains('2_PROSANTE ', case=False)]
df_autre_profil_2prosantesansats = df_prosante[df_prosante['Profile'].str.contains('2_PROSANTE_SANSATS', case=False)]
df_autre_profil_adminprosante = df_prosante[df_prosante['Profile'].str.contains('3_ADMIN_PROSANTE', case=False)]
df_autre_profil_prescripteurs = df_prosante[df_prosante['Profile'].str.contains('5_PRESCRIPTEURSS', case=False)]
df_autre_profil_prosante = df_prosante[df_prosante["Profile"].str.contains('2_PROSANTE|2_PROSANTE_SANSATS', case=False) == False]

# Tableau des comptes crées automatiquement par profession
df_automatiquement_prosante = df_prosante[df_prosante['Cree_Par']=='                                                                                                                                                                                                                                                                ']
df_automatiquement_infirmier = df_infirmier[df_infirmier['Cree_Par']=='                                                                                                                                                                                                                                                                ']
df_automatiquement_biologiste = df_biologiste[df_biologiste['Cree_Par']=='                                                                                                                                                                                                                                                                ']
df_automatiquement_pharmacien = df_pharmacien[df_pharmacien['Cree_Par']=='                                                                                                                                                                                                                                                                ']
df_automatiquement_medecin = df_medecin[df_medecin['Cree_Par']=='                                                                                                                                                                                                                                                                ']
df_automatiquement_dentiste = df_dentiste[df_dentiste['Cree_Par']=='                                                                                                                                                                                                                                                                ']
df_automatiquement_masseur = df_masseur[df_masseur['Cree_Par']=='                                                                                                                                                                                                                                                                ']
df_automatiquement_sage = df_sage[df_sage['Cree_Par']=='                                                                                                                                                                                                                                                                ']
df_automatiquement_MD = df_automatiquement_prosante[df_automatiquement_prosante.LoginName.str.contains('MD', case=False, na=False)]
df_automatiquement_infirmier_MD = df_automatiquement_infirmier[df_automatiquement_infirmier.LoginName.str.contains('MD', case=False, na=False)]
df_automatiquement_pharmacien_MD = df_automatiquement_pharmacien[df_automatiquement_pharmacien.LoginName.str.contains('MD', case=False, na=False)]
df_automatiquement_medecin_MD = df_automatiquement_medecin[df_automatiquement_medecin.LoginName.str.contains('MD', case=False, na=False)]

# Création des fichiers csv
if not os.path.exists(folder + '/CompteParProfession'):
    os.makedirs(folder + '/CompteParProfession')
os.chdir(folder + '/CompteParProfession')
df_manuellement_infirmier.to_csv("compte_manuellement_infirmier.csv", sep=";", index=False)
df_manuellement_pharmacien.to_csv("compte_manuellement_pharmacien.csv", sep=";", index=False)
df_manuellement_medecin.to_csv("compte_manuellement_medecin.csv", sep=";", index=False)
df_manuellement_dentiste.to_csv("compte_manuellement_dentiste.csv", sep=";", index=False)
df_manuellement_masseur.to_csv("compte_manuellement_masseur.csv", sep=";", index=False)
df_manuellement_sage.to_csv("compte_manuellement_sageFemme.csv", sep=";", index=False)
df_manuellement_prosante.to_csv("compte_manuellement_prosante.csv", sep=";", index=False)
df_manuellement_autres.to_csv("compte_manuellement_autres.csv", sep=";", index=False)

# Si on veut avoir les excels de ces professions il suffit de décommenter les lignes suivantes

# df_manuellement_ambulancier.to_csv("compte_manuellement_ambulancier.csv", sep=";", index=False)
# df_manuellement_manipulateur.to_csv("compte_manuellement_manipulateur.csv", sep=";", index=False)
# df_manuellement_opticien.to_csv("compte_manuellement_opticien.csv", sep=";", index=False)
# df_manuellement_orthophoniste.to_csv("compte_manuellement_orthophoniste.csv", sep=";", index=False)
# df_manuellement_orthoptiste.to_csv("compte_manuellement_orthoptiste.csv", sep=";", index=False)
# df_manuellement_osteopathe.to_csv("compte_manuellement_osteopathe.csv", sep=";", index=False)
# df_manuellement_pedicure.to_csv("compte_manuellement_pedicure.csv", sep=";", index=False)
# df_manuellement_podologue.to_csv("compte_manuellement_podologue.csv", sep=";", index=False)
# df_manuellement_prothesiste.to_csv("compte_manuellement_prothesiste.csv", sep=";", index=False)
# df_manuellement_psychiatre.to_csv("compte_manuellement_psychiatre.csv", sep=";", index=False)
df_prosante.to_csv("compte_prosante.csv", sep=";", index=False)
df_autre_profil_prosante.to_csv("autre_profil_que_prosante.csv", sep=";", index=False)
df_infirmier.to_csv("compte_infirmier_total.csv", sep=";", index=False)

df_automatiquement_MD.to_csv("compte_automatiquement_prosante_MD.csv", sep=";", index=False)
df_automatiquement_infirmier_MD.to_csv("compte_automatiquement_infirmier_MD.csv", sep=";", index=False)
os.chdir(folder)

# Actualisation des slides
if not os.path.exists(folder + '/ActualisationSlides'):
    os.makedirs(folder + '/ActualisationSlides')
os.chdir(folder + '/ActualisationSlides')

# Nombre de compte sur SIDEP
df_compte = pd.DataFrame([["Statut", "Nombre"],["ACTIF", nb_actif],["INACTIF", nb_inactif],["Total", nb_total],
                                [""],["Nombre de compte INACTIF sur la production de SIDEP parmi les " + str(nb_actif) + " ACTIF", ''],
                                ["Statut", "Nombre"],["ACTIF et jamais accédé (pas de date d'accès)" , nb_jamais_accede],
                                ["ACTIF et en sommeil (plus d'accès à SIDEP depuis le " + str(troismois.date()) + ")", nb_sommeil],
                                ["Total", nb_inactif_total],[""],["Total ACTIF", nb_actif_total]],
                                columns=['Nombre de compte sur la production de SIDEP', ' '])
csv_compte = df_compte.to_csv('1_nombre_compte_SIDEP.csv', sep=";", index=False, encoding='ISO-8859-1')

# Nombre de compte par organisation
c_orga = pd.concat([c_orga, df_total_orga])
csv_orga = c_orga.to_csv('2_nombre_compte_organisation.csv', sep=";", index_label=['Organisation'], header=['Nombre'], encoding='ISO-8859-1')

# Détail des PROSANTE
tot_prof = (len(df_biologiste)+len(df_pharmacien)+len(df_medecin)+len(df_dentiste)+len(df_masseur)+len(df_sage))
tot_auto = (len(df_automatiquement_biologiste)+len(df_automatiquement_pharmacien)+len(df_automatiquement_medecin)+
            len(df_automatiquement_dentiste)+len(df_automatiquement_masseur)+len(df_automatiquement_sage))
tot_man = (len(df_manuellement_biologiste)+len(df_manuellement_pharmacien)+len(df_manuellement_medecin)+
            len(df_manuellement_dentiste)+len(df_manuellement_masseur)+len(df_manuellement_sage))

df_detail_prosante = pd.DataFrame([["Profession", "Nombre de professionnels", "Nombre de comptes auto-créés", "Comptes crées manuellement"],
                                   ["ADELI (Infirmier)", str(len(df_infirmier)), str(len(df_automatiquement_infirmier)), str(len(df_manuellement_infirmier))],
                                   ["", "", "", ""],
                                   ["Biologiste", str(len(df_biologiste)), str(len(df_automatiquement_biologiste)), str(len(df_manuellement_biologiste))],
                                   ["Chirurgien Dentiste", str(len(df_dentiste)), str(len(df_automatiquement_dentiste)), str(len(df_manuellement_dentiste))],
                                   ["Masseur Kinésithérapeute", str(len(df_masseur)), str(len(df_automatiquement_masseur)), str(len(df_manuellement_masseur))],
                                   ["Médecin", str(len(df_medecin)), str(len(df_automatiquement_medecin)), str(len(df_manuellement_medecin))],
                                   ["Pharmacien", str(len(df_pharmacien)), str(len(df_automatiquement_pharmacien)), str(len(df_manuellement_pharmacien))],
                                   ["Sage femme", str(len(df_sage)), str(len(df_automatiquement_sage)), str(len(df_manuellement_sage))],
                                   ["Total général", str(tot_prof), str(tot_auto), str(tot_man)],
                                   ["", "", "", ""],
                                   ["Autres comptes RPPS", str(len(df_autres_total)), "", str(len(df_manuellement_autres))],
                                   ["", "", "", ""],
                                   ["Total organisation PROSANTE", "", "", str(len(df_prosante))],
                                   ["", "", "", ""],
                                   ["Profils dans l'organisation PROSANTE (déjà pris en compte dans le comptage, colonne 'Profile')", "", "", ""],
                                   ["2_PROSANTE", str(len(df_autre_profil_2prosante)), "", ""],
                                   ["2_PROSANTE_SANSATS", str(len(df_autre_profil_2prosantesansats)), "", ""],
                                   ["3_ADMIN_PROSANTE", str(len(df_autre_profil_adminprosante)), "", ""],
                                   ["5_PRESCRIPTEURSS", str(len(df_autre_profil_prescripteurs)), "", ""],
                                   ["Autres profils que PROSANTE (différents de 2_PROSANTE et de 2_PROSANTE_SANSATS)", str(len(df_autre_profil_prosante)), "", ""]])

csv_PROSANTE = df_detail_prosante.to_csv('3_detail_PROSANTE.csv', sep=";", index=False, header=False, encoding='ISO-8859-1')

# Tableau autres comptes PROSANTE créés manuellement
tot_autre = (len(df_manuellement_ambulancier)+len(df_manuellement_manipulateur)+len(df_manuellement_opticien)+
            len(df_manuellement_orthophoniste)+len(df_manuellement_orthoptiste)+len(df_manuellement_osteopathe)+
            len(df_manuellement_pedicure)+len(df_manuellement_podologue)+len(df_manuellement_prothesiste)+
            len(df_manuellement_psychiatre)+len(df_manuellement_autres_X))
df_detail_autre_prosante = pd.DataFrame([["Profession", "Nombre de comptes"],
                                   ["Ambulancier", str(len(df_manuellement_ambulancier))],
                                   ["Manipulateur", str(len(df_manuellement_manipulateur))],
                                   ["Opticien", str(len(df_manuellement_opticien))],
                                   ["Orthophoniste", str(len(df_manuellement_orthophoniste))],
                                   ["Orthoptiste", str(len(df_manuellement_orthoptiste))],
                                   ["Ostéopathe", str(len(df_manuellement_osteopathe))],
                                   ["Pédicure-Podologue", str(len(df_manuellement_pedicure))],
                                   ["Podologue", str(len(df_manuellement_podologue))],
                                   ["Prothésiste", str(len(df_manuellement_prothesiste))],
                                   ["Psychiatre", str(len(df_manuellement_psychiatre))],
                                   ["Autres comptes (non défini)", str(len(df_manuellement_autres_X))],
                                   ["", ""],
                                   ["Total général", str(tot_autre)]])

csv_autre_PROSANTE = df_detail_autre_prosante.to_csv('4_detail_autre_PROSANTE.csv', sep=";", index=False, header=False, encoding='ISO-8859-1')

# Comptes créés par le formulaire
# Comptes actifs pour les infirmiers, pharmaciens et médecins
# On extrait le numéro ADELI/RPPS du login et du nom de groupe et on les compare à celui de la liste complète du formulaire
df_adeli_rpps_MD = df_automatiquement_infirmier_MD.assign(ADELI_RPPS_LOGIN=0)
df_adeli_rpps_MD = df_adeli_rpps_MD.assign(ADELI_RPPS_GROUPE=0)
nb_inf_form = 0
for i in df_automatiquement_infirmier_MD.index:
    nom_groupe = df_automatiquement_infirmier_MD['NomGroupe'][i]
    split_nom_groupe = nom_groupe.split('_')
    df_adeli_rpps_MD['ADELI_RPPS_GROUPE'][i] = split_nom_groupe[1]
    login = df_automatiquement_infirmier_MD['LoginName'][i]
    split = login.split(' ')
    login2 = split[0][4:]
    df_adeli_rpps_MD['ADELI_RPPS_LOGIN'][i] = str(login2)
    for k in df2.index:
        if (split_nom_groupe[1]==df2['ADELI_RPPS'][k] or login2==df2['ADELI_RPPS'][k]):
            nb_inf_form +=1
# Boucle ajouté car compte créé manuellement avant la mise en place du script
for i in df_manuellement_infirmier.index:
    if (df_manuellement_infirmier['Cree_Par'][i]=='7041477                                                                                                                                                                                                                                                         '):
        nb_inf_form +=1
        
# Comptes actifs et inactifs créés par le formulaire pour les 3 mêmes professions
df_infirmier_tot = df[df.NomGroupe.str.contains('infirmier', case=False)]
df_pharmacien_tot = df[df.NomGroupe.str.contains('pharmacien', case=False)]
df_medecin_tot = df[df.NomGroupe.str.contains('medecin', case=False)]
df_adeli_rpps_MD = df_infirmier_tot.assign(ADELI_RPPS_LOGIN=0)
df_adeli_rpps_MD = df_adeli_rpps_MD.assign(ADELI_RPPS_GROUPE=0)
nb_inf_form_tot = 0
for i in df_infirmier_tot.index:
    nom_groupe = df_infirmier_tot['NomGroupe'][i]
    split_nom_groupe = nom_groupe.split('_')
    df_adeli_rpps_MD['ADELI_RPPS_GROUPE'][i] = split_nom_groupe[1]
    login = df_infirmier_tot['LoginName'][i]
    # Le Cree_Par 7041477 correspond aux comptes créés manuellement via le formulaire
    if(str(login).startswith('MD_') and (df_infirmier_tot['Cree_Par'][i]=='                                                                                                                                                                                                                                                                ' 
                                         or df_infirmier_tot['Cree_Par'][i]=='7041477                                                                                                                                                                                                                                                         ')):
        split = login.split(' ')
        login2 = split[0][4:]
        df_adeli_rpps_MD['ADELI_RPPS_LOGIN'][i] = str(login2)
        for k in df2.index:
            if (split_nom_groupe[1]==df2['ADELI_RPPS'][k] or login2==df2['ADELI_RPPS'][k]):
                nb_inf_form_tot +=1
                    
df_compte_form = pd.DataFrame([["Profession", "Nombre de comptes actif", "Nombre de comptes total"],
                                   ["Infirmier", str(nb_inf_form), str(nb_inf_form_tot)],
                                   ["Total général", str(nb_inf_form), str(nb_inf_form_tot)]])

csv_compte_form = df_compte_form.to_csv('5_detail_compte_formulaire.csv', sep=";", index=False, header=False, encoding='ISO-8859-1')
os.chdir(folder)

# Traitement des doublons 
if not os.path.exists(folder + '/Doublons'):
    os.makedirs(folder + '/Doublons')
os.chdir(folder + '/Doublons')
# pour les infirmiers
df_adeli_rpps = df_infirmier.assign(ADELI_RPPS_GROUPE=0)
df_adeli_rpps['Prenom'] = df_adeli_rpps['Prenom'].str.upper()
df_adeli_rpps['Nom'] = df_adeli_rpps['Nom'].str.upper()

df_adeli_rpps = df_adeli_rpps.assign(ADELI_LOGIN=0)

for k in df_infirmier.index:
    nom_groupe = df_infirmier['NomGroupe'][k]
    split_nom_groupe = nom_groupe.split('_')
    df_adeli_rpps['ADELI_RPPS_GROUPE'][k] = split_nom_groupe[1]
    login = df_infirmier['LoginName'][k]
    if(str(login).startswith('MD_')):
        split_login = login.split('_')
        split2 = split_login[1].split(' ')
        df_adeli_rpps['ADELI_LOGIN'][k] = str(split2[0])
    else:
        split_login = str(login).split(' ')
        df_adeli_rpps['ADELI_LOGIN'][k] = str(split_login[0])
        
doublons_inf = df_adeli_rpps[df_adeli_rpps.duplicated(subset=['ADELI_RPPS_GROUPE', 'Prenom'], keep=False)]
doublons2_inf = df_adeli_rpps[df_adeli_rpps.duplicated(subset=['ADELI_LOGIN'], keep=False)]

doublons_inf_vf = pd.concat([doublons_inf,doublons2_inf])
doublons_inf_vf = doublons_inf_vf.drop_duplicates()
doublons_inf_vf.to_excel("liste_doublons_infirmier.xlsx", index=False)
    
# Pour les pharmaciens
df_adeli_rpps = df_pharmacien.assign(ADELI_RPPS_GROUPE=0)
df_adeli_rpps['Prenom'] = df_adeli_rpps['Prenom'].str.upper()
df_adeli_rpps['Nom'] = df_adeli_rpps['Nom'].str.upper()
df_adeli_rpps = df_adeli_rpps.assign(ADELI_LOGIN=0)

for k in df_pharmacien.index:
    nom_groupe = df_pharmacien['NomGroupe'][k]
    split_nom_groupe = nom_groupe.split('_')
    df_adeli_rpps['ADELI_RPPS_GROUPE'][k] = split_nom_groupe[1]
    login = df_pharmacien['LoginName'][k]
    if(str(login).startswith('MD_')):
        split_login = login.split('_')
        split2 = split_login[1].split(' ')
        df_adeli_rpps['ADELI_LOGIN'][k] = str(split2[0])
    else:
        split_login = str(login).split(' ')
        df_adeli_rpps['ADELI_LOGIN'][k] = str(split_login[0])
        
doublons_phar = df_adeli_rpps[df_adeli_rpps.duplicated(subset=['ADELI_RPPS_GROUPE', 'Prenom'], keep=False)]
doublons2_phar = df_adeli_rpps[df_adeli_rpps.duplicated(subset=['ADELI_LOGIN'], keep=False)]

doublons_phar_vf = pd.concat([doublons_phar,doublons2_phar])
doublons_phar_vf = doublons_phar_vf.drop_duplicates()
doublons_phar_vf.to_excel("liste_doublons_pharmacien.xlsx", index=False)

# Doublons autres
df_adeli_rpps = df_prosante.assign(ADELI_RPPS_GROUPE=0)
df_adeli_rpps['Prenom'] = df_adeli_rpps['Prenom'].str.upper()
df_adeli_rpps['Nom'] = df_adeli_rpps['Nom'].str.upper()
df_adeli_rpps = df_adeli_rpps.assign(ADELI_LOGIN=0)

for k in df_prosante.index:
    login = df_prosante['LoginName'][k]
    if(str(login).startswith('MD_')):
        split_login = login.split('_')
        split2 = split_login[1].split(' ')
        df_adeli_rpps['ADELI_LOGIN'][k] = str(split2[0])
    else:
        split_login = str(login).split(' ')
        df_adeli_rpps['ADELI_LOGIN'][k] = str(split_login[0])
        
doublons_autres = df_adeli_rpps[df_adeli_rpps.duplicated(subset=['ADELI_LOGIN'], keep=False)]

df_doublons_inf_phar = pd.concat([doublons_inf_vf,doublons_phar_vf])
df_doublons_autres_X = doublons_autres.copy()

for i in df_doublons_inf_phar.index:
    for k in df_doublons_autres_X.index:
        if(k==i):
            df_doublons_autres_X = df_doublons_autres_X.drop(k)
df_doublons_autres_X.to_excel("liste_doublons_autres.xlsx", index=False)
os.chdir(folder)
