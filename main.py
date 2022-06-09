import os
import numpy as np
import xlrd
from pathlib import Path
import tkinter as tk
from tkinter import ttk

import matplotlib
from matplotlib import dates

matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.figure import Figure

import datetime as dt
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from scipy.optimize import curve_fit
import matplotlib.ticker as ticker

from PIL import Image, ImageTk

#  -------------------------------------------------------------------------------------------------- Chemin d'accès  --------------------------------------------------------------------------------------------------
# Récupération de tous les chemins d'acces de base (là où il y a les dossiers Pv Sable et Pv Flexion
path_general = sorted(Path("F:\Boulot\Ferry_Capitain\Synthèse controle sable").iterdir(), key=os.path.getmtime) # A CHANGER EN FONCTION DU RESEAU, ici ma clé USB fonctionne 17/02/2021
# path_general = sorted(Path("E:\Sablerie FERRY\Programme").iterdir(), key=os.path.getmtime) # A CHANGER EN FONCTION DU RESEAU, ici ma clé USB fonctionne 17/02/2021
# path_general = sorted(Path("//192.168.200.109/LaboSable").iterdir(), key=os.path.getmtime)  # A CHANGER EN FONCTION DU RESEAU

path_sous_general_sable = []
print(path_general, "general")  # Ok 23/02/2021
paths = []  # liste finale sable
paths_flexion = []  # liste finale flexion
for i in range(len(path_general)):  # on va chercher tous les sous dossiers du chemin donné
    if str(path_general[i])[-20:].find('Pv Sable') >= 0:
        path_sous_general_sable = sorted(Path(path_general[i]).iterdir(), key=os.path.getmtime)
        # print(path_sous_general_sable, "sous dossier sable") #Ok 23/02/2021
        for j in range(len(path_sous_general_sable)):
            if str(path_sous_general_sable[j])[-20:].find('Analyses') >= 0 or str(path_sous_general_sable[j])[
                                                                              -20:].find('analyses') >= 0:
                Temp = sorted(Path(path_sous_general_sable[j]).iterdir(), key=os.path.getmtime)
                for k in range(len(Temp)):
                    # print('ok') # pour la visualitation temporelle
                    paths.append(Temp[k])
                # print(paths, str(path_sous_general_sable[j])[-20:]) #fonctionne 23/02/2021
    if str(path_general[i])[-20:].find('Pv Flexion') >= 0:
        path_sous_general_flexion = sorted(Path(path_general[i]).iterdir(), key=os.path.getmtime)
        # print(path_sous_general_flexion, "flexion")
        for j in range(len(path_sous_general_flexion)):
            if str(path_sous_general_flexion[j])[-20:].find('2011') >= 0 or str(path_sous_general_flexion[j])[
                                                                            -20:].find('2012') >= 0 or str(
                    path_sous_general_flexion[j])[-20:].find('2013') >= 0 or str(path_sous_general_flexion[j])[
                                                                             -20:].find('2014') >= 0 or str(
                    path_sous_general_flexion[j])[-20:].find('2015') >= 0 or str(path_sous_general_flexion[j])[
                                                                             -20:].find('2016') >= 0 or str(
                    path_sous_general_flexion[j])[-20:].find('2017') >= 0 or str(path_sous_general_flexion[j])[
                                                                             -20:].find('2018') >= 0 or str(
                    path_sous_general_flexion[j])[-20:].find('2019') >= 0 or str(path_sous_general_flexion[j])[
                                                                             -20:].find('2020') >= 0 or str(
                    path_sous_general_flexion[j])[-20:].find('2021') >= 0 or str(path_sous_general_flexion[j])[
                                                                             -20:].find('2022') >= 0 or str(
                    path_sous_general_flexion[j])[-20:].find('2023') >= 0 or str(path_sous_general_flexion[j])[
                                                                             -20:].find('2024') >= 0 or str(
                    path_sous_general_flexion[j])[-20:].find('2025') >= 0 or str(path_sous_general_flexion[j])[
                                                                             -20:].find('2026') >= 0 or str(
                    path_sous_general_flexion[j])[-20:].find('2027') >= 0 or str(path_sous_general_flexion[j])[
                                                                             -20:].find('2028') >= 0 or str(
                    path_sous_general_flexion[j])[-20:].find('2029') >= 0 or str(path_sous_general_flexion[j])[
                                                                             -20:].find('2030') >= 0:
                Temp = sorted(Path(path_sous_general_flexion[j]).iterdir(), key=os.path.getmtime)
                for k in range(len(Temp)):
                    paths_flexion.append(Temp[k])
    print(i, '/', len(path_general), 'lecture des sous dossiers')  # avancement

# print(paths, "sable") #vérif de la commande - fonctionne 23/02/2021
# print(paths_flexion[0], "flexion") #vérif de la commande - fonctionne 23/02/2021
print("fin de la lecture de tous les excels du Labo Sable")


#  -------------------------------------------------------------------------------------------------- Fonction Refus  --------------------------------------------------------------------------------------------------
# Fonction qui renvoi à partir du chemin d'acces à un excel la liste des refus des tamis de 1 à 12 dans une liste R simple
# R = [date, R tamis 1, R tamis 2, ..., R tamis 12, Indice AFS, pH, Poussières, perte au feu]
def Refus(path):
    classeur = xlrd.open_workbook(path)
    nom_des_feuilles = classeur.sheet_names()
    feuille = classeur.sheet_by_name(nom_des_feuilles[0])
    R = []
    R.append(feuille.cell(1, 10).value)  # date
    for i in range(12):  # tout les refus des tamis de 1 à 12
        R.append(feuille.cell(11 + i, 3).value)
    # print(feuille.cell(11,3)) #vérif (premier refus)
    R.append(feuille.cell(9, 8).value)  # indice AFS
    R.append(feuille.cell(9, 11).value)  # pH
    R.append(feuille.cell(17, 8).value)  # poussières
    R.append(feuille.cell(17, 11).value)  # perte au feu

    return R


# print(Refus(paths[0])) # vérif - fonctionne

#  -------------------------------------------------------------------------------------------------- Fonction Flexion  --------------------------------------------------------------------------------------------------
# Fonction qui renvoi à partir du chemin d'acces à un excel la liste des flexions AVEC l'appartenance à un chaniter (et une machine ?)
def Flexion(path):
    classeur = xlrd.open_workbook(path)
    nom_des_feuilles = classeur.sheet_names()
    F_MA = classeur.sheet_by_name(nom_des_feuilles[0])  # feuille MA
    F_F1 = classeur.sheet_by_name(nom_des_feuilles[1])  # feuille F1
    F_F3 = classeur.sheet_by_name(nom_des_feuilles[2])  # feuille F3
    F = []
    # print(path)
    # print(str(path)[-6:-4])
    # print('ok', str(path).find('Flexion')+8)
    # print(str(path)[str(path).find('Flexion')+8:str(path).find('Flexion')+12])

    # print(dt.datetime.strptime(str(path)[-6:-4]+'-Monday-'+str(path)[str(path).find('Flexion')+8:str(path).find('Flexion')+12], '%V-%A-%G')) #fonctione 08/02/2021
    # ici on utilise le numéro de semaine du nom de l'excel ainsi que le nom du sous dossier dans lequel il est contenu pour récupéré la date
    # en effet les dates inscrites dans les documents ne sont pas uniformes dans le temps...
    F.append(dt.datetime.strptime(
        str(path)[-6:-4] + '-Monday-' + str(path)[str(path).find('Flexion') + 8:str(path).find('Flexion') + 12],
        '%V-%A-%G'))  # date
    # print(F[0]) # test date - fonctionne 10/02/21
    if int(str(path)[str(path).find('Flexion') + 8:str(path).find(
            'Flexion') + 12]) > 2016:  # pour les fichiers datant de 2016 ou + :
        for j in [F_MA, F_F1, F_F3]:
            for i in [17, 19, 21, 30, 32, 34, 43, 45, 47]:
                # print(i, j.cell(i,2).value) # test lecture des flexions 2016 + - fonctionne 10/02/21
                if type(j.cell(i, 2).value) != str:
                    Moy = 0
                    nb = 0
                    for k in [2, 3, 4, 5, 6, 7, 8, 9]:
                        if j.cell(i, k).value != '':
                            Moy = Moy + int(j.cell(i, k).value)  # /10 divise ici par 10 pour être en Mpa
                            nb = nb + 1
                        else:
                            break
                    F.append(Moy / nb)
                else:
                    F.append(np.nan)
    else:  # pour les fichiers antérieurs à 2016 (changement de lignes)
        for j in [F_MA, F_F1, F_F3]:
            for i in [16, 18, 20, 29, 31, 33, 42, 44, 46]:
                # print(i, j.cell(i,2).value) # test lecture des flexions 2016 - -fonctionne 10/02/21
                if type(j.cell(i, 2).value) != str:
                    Moy = 0
                    nb = 0
                    for k in [2, 3, 4, 5, 6, 7, 8, 9]:
                        if j.cell(i, k).value != '':
                            Moy = Moy + int(j.cell(i, k).value)  # /10 divise ici par 10 pour être en Mpa
                            nb = nb + 1
                        else:
                            break
                    F.append(Moy / nb)
                else:
                    F.append(np.nan)
    return F  # voir note sur i pad


# print(Flexion(paths_flexion[0]))# test de la fonction Flexion - fonctionne 10/02/21


# SF est la grande liste des flexions
SF = []
for i in range(len(paths_flexion)):
    if str(paths_flexion[i])[-6:-5] == '0' or str(paths_flexion[i])[-6:-5] == '1' or str(paths_flexion[i])[
                                                                                     -6:-5] == '2' or str(
            paths_flexion[i])[-6:-5] == '3' or str(paths_flexion[i])[-6:-5] == '4' or str(paths_flexion[i])[
                                                                                      -6:-5] == '5' or str(
            paths_flexion[i])[-6:-5] == '6' or str(paths_flexion[i])[-6:-5] == '7' or str(paths_flexion[i])[
                                                                                      -6:-5] == '8' or str(
            paths_flexion[i])[-6:-5] == '9':
        SF.append(Flexion(paths_flexion[i]))
    else:
        continue
    print(i, '/', len(paths_flexion), 'traitement des Excels des Flexions')
# print(SF) # fonctionne 23/02/2021

print('fin du traitement des Flexions')

#  -------------------------------------------------------------------------------------------------- Regroupe les informations des Refus en les catégorisans--------------------------------------------------------------------------------------------------
# S est la grande liste des sables
# S = [sable récup[    [silo[  MA[Refus[]],F1[Refus[]],F3[Refus[]],BMM[Refus[]]    ], Sablerie[    MA[Refus[]],F1[Refus[]],F3[Refus[]],BMM[Refus[]]    ] etc
S = [[[[], [], [], []], [[], [], [], []], [[], [], [], []], [[], [], [], []], [[], [], [], []]],
     [[[], [], [], []], [[], [], [], []], [[], [], [], []], [[], [], [], []], [[], [], [], []]],
     [[[], [], [], []], [[], [], [], []], [[], [], [], []], [[], [], [], []], [[], [], [], []]],
     [[[], [], [], []], [[], [], [], []], [[], [], [], []], [[], [], [], []], [[], [], [], []]],
     [[[], [], [], []], [[], [], [], []], [[], [], [], []], [[], [], [], []], [[], [], [], []]]]
for i in range(len(paths)):
    # reset des indexs
    Type = -1  # Type = 0 : sable récup / Type = 1 : sable RTH / Type = 2 : silice neuve / Type = 3 : chromite neuve / Type = 4 : Vrac / Type = -1 : Erreur nom fichier
    Prelev = -1  # Prelev = 0 : silo / Prelev = 1 : sablerie / Prelev = 2 : malaxeur / Prelev = 3 : Big Bag / Prelev = 4 : Vrac / Prelev = -1 : Erreur nom fichier
    Chantier = -1  # Chantier = 0 : MA / Chantier = 1 : F1 / Chantier = 2 : F3 / Chantier = 3 : BMM / Chantier = -1 : Erreur nom fichier
    if (str(paths[i])[-40:].find('sable récup') >= 0 or str(paths[0])[-40:].find('sable recup') >= 0):
        Type = 0

        if str(paths[i])[-40:].find('silo') >= 0:
            Prelev = 0
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('sablerie') >= 0:
            Prelev = 1
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('malaxeur') >= 0:
            Prelev = 2
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('BB') >= 0:
            Prelev = 3
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('vrac') >= 0:
            Prelev = 4
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))

    elif str(paths[i])[-40:].find('sable RTH') >= 0:
        Type = 1

        if str(paths[i])[-40:].find('silo') >= 0:
            Prelev = 0
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('sablerie') >= 0:
            Prelev = 1
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('malaxeur') >= 0:
            Prelev = 2
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('BB') >= 0:
            Prelev = 3
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('vrac') >= 0:
            Prelev = 4
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))

    elif str(paths[i])[-40:].find('silice neuve') >= 0:
        Type = 2
        S[Type][0][0].append(Refus(paths[i]))  # les sables neuf n'ont pas de prelevement ou de chanitier referencé


    elif str(paths[i])[-40:].find('chromite neuve') >= 0:
        Type = 3
        S[Type][0][0].append(Refus(paths[i]))  # les sables neuf n'ont pas de prelevement ou de chanitier referencé

    elif (str(paths[i])[-40:].find('chromite recup') >= 0 or str(paths[i])[-40:].find('chromite récup') >= 0 or str(
            paths[i])[-40:].find('Chromite récup') >= 0):
        Type = 4

        if str(paths[i])[-40:].find('silo') >= 0:
            Prelev = 0
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('sablerie') >= 0:
            Prelev = 1
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('malaxeur') >= 0:
            Prelev = 2
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('BB') >= 0:
            Prelev = 3
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
        elif str(paths[i])[-40:].find('vrac') >= 0:
            Prelev = 4
            if str(paths[i])[-40:].find('MA') >= 0:
                Chantier = 0
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F1') >= 0:
                Chantier = 1
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('F3') >= 0:
                Chantier = 2
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
            elif str(paths[i])[-40:].find('BMM') >= 0:
                Chantier = 3
                S[Type][Prelev][Chantier].append(Refus(paths[i]))
    print(i, '/', len(paths), 'traitement des tests classiques (hors flexion)')

# print(S)
print("fin du traitement des tests classiques (hors flexion)")

#  -------------------------------------------------------------------------------------------------- Début tkinter (interface)  --------------------------------------------------------------------------------------------------
window = tk.Tk()
window.title("Synthèse sable")

window.columnconfigure([0, 1], weight=1, minsize=75)
window.rowconfigure([0], weight=1)
window.rowconfigure([1], weight=10, minsize=50)

# Frame du choix des dates ----------------------------------------------------
frm_date = tk.Frame(master=window)
frm_date.grid(row=0, column=0, sticky="nsew")

frm_date.columnconfigure([0, 1, 2, 3], weight=1, minsize=75)
frm_date.rowconfigure([0, 1, 2, 3], weight=1, minsize=50)

lb_index_jours = tk.Label(master=frm_date, text="Jour").grid(row=1, column=1)
lb_index_mois = tk.Label(master=frm_date, text="Mois").grid(row=1, column=2)
lb_index_ans = tk.Label(master=frm_date, text="Ans").grid(row=1, column=3)
Jours_start = ttk.Combobox(master=frm_date,
                           value=["", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15",
                                  "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29",
                                  "30", "31"])
Mois_start = ttk.Combobox(master=frm_date,
                          value=["", "1 - Janvier", "2 - Février", "3 - Mars", "4 - Avril", "5 - Mai", "6 - Juin",
                                 "7 - Juillet", "8 - Aout", "9 - Septembre", "10 - Octobre", "11 - Novembre",
                                 "12 - Décembre"])
Ans_start = ttk.Combobox(master=frm_date,
                         value=["", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2023", "2024",
                                "2025", "2026", "2027", "2028", "2029", "2030"])
lb_date_start = tk.Label(master=frm_date, text=" Date début :")
Jours_end = ttk.Combobox(master=frm_date,
                         value=["", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15",
                                "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29",
                                "30", "31"])
Mois_end = ttk.Combobox(master=frm_date,
                        value=["", "1 - Janvier", "2 - Février", "3 - Mars", "4 - Avril", "5 - Mai", "6 - Juin",
                               "7 - Juillet", "8 - Aout", "9 - Septembre", "10 - Octobre", "11 - Novembre",
                               "12 - Décembre"])
Ans_end = ttk.Combobox(master=frm_date,
                       value=["", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2023", "2024",
                              "2025", "2026", "2027", "2028", "2029", "2030"])
lb_date_end = tk.Label(master=frm_date, text=" Date fin :")

lb_date_start.grid(row=2, column=0, sticky="w")
Jours_start.grid(row=2, column=1, padx=5)
Mois_start.grid(row=2, column=2, padx=5)
Ans_start.grid(row=2, column=3, padx=5)
lb_date_end.grid(row=3, column=0, sticky="w")
Jours_end.grid(row=3, column=1, padx=5)
Mois_end.grid(row=3, column=2, padx=5)
Ans_end.grid(row=3, column=3, padx=5)

# Logo dans la frame date
im = Image.open('logo_FC.png', 'r')
#im = im.resize((150, 150), Image.LANCZOS)
im_logo = ImageTk.PhotoImage(im, master=frm_date)
logo = tk.Canvas(master=frm_date)  # , width = 100, height = 100)
FC = logo.create_image(200, 150, image=im_logo)  # , anchor = nsew
logo.grid(row=0, column=0)

# Texte auteur
lb_auteur = tk.Label(master=frm_date,
                     text="Outil de synthèse des données\n du contrôle des sables\n\n -- Ferry Capitain --",
                     font=("Calibri", 24)).grid(row=0, column=1, columnspan=3)

# Frame de tous les choix --------------------------------------------------------------

frm_choix = tk.Frame(master=window, relief='groove', bd=5)
frm_choix.grid(row=0, column=1, sticky="nsew")

frm_choix.columnconfigure([0, 1, 2, 3, 4], weight=1, minsize=75)
frm_choix.rowconfigure([0, 1, 2, 3, 4, 5, 6, 7, 8], weight=1, minsize=50)

lb_sable = tk.Label(frm_choix, text="Sable : ")
lb_prelev = tk.Label(frm_choix, text="Prélévement : ")
lb_chantier = tk.Label(frm_choix, text="Chantier : ")
lb_courbe = tk.Label(frm_choix, text="Ordonnées :")
lb_sable.grid(row=0, column=0)
lb_prelev.grid(row=0, column=1)
lb_chantier.grid(row=0, column=2)
lb_courbe.grid(row=0, column=3)


# Sable

def Neuve():  # commande lorsque le salbe neuf est sélectioné
    if Sable[2].get() == 1 or Sable[3].get() == 1:
        k = 0
        for i in [check_silo, check_sablerie, check_malaxeur, check_BB, check_vrac]:
            i.state(['disabled', '!selected'])
            Prelev[k].set(0)
            k = k + 1
        k = 0
        for i in [check_MA, check_F1, check_F3]:
            i.state(['disabled', '!selected'])
            Chantier[k].set(0)
            k = k + 1
        Prelev[0].set(1)
        check_silo.state(['!selected'])  # je cache ma supércherie pour éviter les malentendu (sinon case grisé coché)
        Chantier[0].set(
            1)  # j'ai rangé les données dans silo MA mais je l'affiche pas en légende à la fin pour rester dans le cas général lors du plot
        check_MA.state(['!selected'])
    else:
        for i in [check_silo, check_sablerie, check_malaxeur, check_BB, check_vrac]:
            i.state(['!disabled'])
        for i in [check_MA, check_F1, check_F3]:
            i.state(['!disabled'])
        Prelev[0].set(0)
        Chantier[0].set(0)
    return


def G_sable():
    B_ord = [check_sable_recup, check_sable_RTH, check_silice_neuve, check_chromite_neuve, check_chromite_recup]
    k = 0
    for i in range(len(Sable)):
        k = k + Sable[i].get()
    if k == 0:
        for i in range(len(Sable)):  # Si aucun n'est coché
            B_ord[i].state(['!disabled'])  # tout reste actif
        return
    elif k == 1:
        for i in range(len(Sable)):  # Si un est coché
            if Sable[i].get() == 0:
                B_ord[i].state(['disabled', '!selected'])
    return


Sable = [tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar()]
Sable_index = ["sable récup", "sable RTH", "silice neuve", "chromite neuve", "chromite récup"]
check_sable_recup = ttk.Checkbutton(frm_choix, text=Sable_index[0], variable=Sable[0], command=G_sable)
check_sable_RTH = ttk.Checkbutton(frm_choix, text=Sable_index[1], variable=Sable[1], command=G_sable)
check_silice_neuve = ttk.Checkbutton(frm_choix, text=Sable_index[2], variable=Sable[2],
                                     command=lambda: [Neuve(), G_sable()])
check_chromite_neuve = ttk.Checkbutton(frm_choix, text=Sable_index[3], variable=Sable[3],
                                       command=lambda: [Neuve(), G_sable()])
check_chromite_recup = ttk.Checkbutton(frm_choix, text=Sable_index[4], variable=Sable[4], command=G_sable)

check_sable_recup.grid(row=1, column=0)
check_sable_RTH.grid(row=2, column=0)
check_silice_neuve.grid(row=3, column=0)
check_chromite_neuve.grid(row=4, column=0)
check_chromite_recup.grid(row=5, column=0)


# Prélevement

def G_prelev():
    B_ord = [check_silo, check_sablerie, check_malaxeur, check_BB, check_vrac]
    k = 0
    for i in range(len(Prelev)):
        k = k + Prelev[i].get()
    if k == 0:
        for i in range(len(Prelev)):  # Si aucun n'est coché
            B_ord[i].state(['!disabled'])  # tout reste actif
        return
    elif k == 1:
        for i in range(len(Prelev)):  # Si un est coché
            if Prelev[i].get() == 0:
                B_ord[i].state(['disabled', '!selected'])
    return


Prelev = [tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar()]
Prelev_index = ["silo", "sablerie", "malaxeur", "Big Bag", "Vrac"]
check_silo = ttk.Checkbutton(frm_choix, text=Prelev_index[0], variable=Prelev[0], command=G_prelev)
check_sablerie = ttk.Checkbutton(frm_choix, text=Prelev_index[1], variable=Prelev[1], command=G_prelev)
check_malaxeur = ttk.Checkbutton(frm_choix, text=Prelev_index[2], variable=Prelev[2], command=G_prelev)
check_BB = ttk.Checkbutton(frm_choix, text=Prelev_index[3], variable=Prelev[3], command=G_prelev)
check_vrac = ttk.Checkbutton(frm_choix, text=Prelev_index[4], variable=Prelev[4], command=G_prelev)

check_silo.grid(row=1, column=1)
check_sablerie.grid(row=2, column=1)
check_malaxeur.grid(row=3, column=1)
check_BB.grid(row=4, column=1)
check_vrac.grid(row=5, column=1)


# Chantier
# Fonction suivant l'event du bouton MA coché

def G_chantier():
    B_ord = [check_MA, check_F1, check_F3]
    k = 0
    for i in range(len(Chantier)):
        k = k + Chantier[i].get()
    if k == 0:
        for i in range(len(Chantier)):  # Si aucun n'est coché
            B_ord[i].state(['!disabled'])  # tout reste actif
        return
    elif k == 1:
        for i in range(len(Chantier)):  # Si un est coché
            if Chantier[i].get() == 0:
                B_ord[i].state(['disabled', '!selected'])
    return


def MA():
    if Courbe[5].get() == 1:
        if Chantier[0].get() == 1:
            check_IMF_F1.state(['disabled', '!selected'])
            Flexion_malaxeur[3].set(0)
            check_SAPIC.state(['disabled', '!selected'])
            Flexion_malaxeur[4].set(0)
            check_FAT_F3.state(['disabled', '!selected'])
            Flexion_malaxeur[5].set(0)
            check_FAT_F5.state(['disabled', '!selected'])
            Flexion_malaxeur[6].set(0)
            check_MA2.state(['!disabled'])
            check_IMF_MA.state(['!disabled'])
            check_FAT.state(['!disabled'])
        elif Chantier[0].get() == 0:
            check_MA2.state(['disabled', '!selected'])
            Flexion_malaxeur[0].set(0)
            check_IMF_MA.state(['disabled', '!selected'])
            Flexion_malaxeur[1].set(0)
            check_FAT.state(['disabled', '!selected'])
            Flexion_malaxeur[2].set(0)
    return


# Fonction suivant l'event du bouton F1 coché
def F1():
    if Courbe[5].get() == 1:
        if Chantier[1].get() == 1:
            check_IMF_F1.state(['!disabled'])
            check_SAPIC.state(['!disabled'])
            check_FAT_F3.state(['disabled', '!selected'])
            Flexion_malaxeur[5].set(0)
            check_FAT_F5.state(['disabled', '!selected'])
            Flexion_malaxeur[6].set(0)
            check_MA2.state(['disabled', '!selected'])
            Flexion_malaxeur[0].set(0)
            check_IMF_MA.state(['disabled', '!selected'])
            Flexion_malaxeur[1].set(0)
            check_FAT.state(['disabled', '!selected'])
            Flexion_malaxeur[2].set(0)
        elif Chantier[1].get() == 0:
            check_IMF_F1.state(['disabled', '!selected'])
            Flexion_malaxeur[3].set(0)
            check_SAPIC.state(['disabled', '!selected'])
            Flexion_malaxeur[4].set(0)
    return


# Fonction suivant l'event du bouton F3 coché
def F3():
    if Courbe[5].get() == 1:
        if Chantier[2].get() == 1:
            check_IMF_F1.state(['disabled', '!selected'])
            Flexion_malaxeur[3].set(0)
            check_SAPIC.state(['disabled', '!selected'])
            Flexion_malaxeur[4].set(0)
            check_FAT_F3.state(['!disabled'])
            check_FAT_F5.state(['!disabled'])
            check_MA2.state(['disabled', '!selected'])
            Flexion_malaxeur[0].set(0)
            check_IMF_MA.state(['disabled', '!selected'])
            Flexion_malaxeur[1].set(0)
            check_FAT.state(['disabled', '!selected'])
            Flexion_malaxeur[2].set(0)
        elif Chantier[2].get() == 0:
            check_FAT_F3.state(['disabled', '!selected'])
            Flexion_malaxeur[5].set(0)
            check_FAT_F5.state(['disabled', '!selected'])
            Flexion_malaxeur[6].set(0)
    return


Chantier = [tk.IntVar(), tk.IntVar(), tk.IntVar()]
Chantier_index = ["MA", "F1", "F3"]
check_MA = ttk.Checkbutton(frm_choix, text=Chantier_index[0], variable=Chantier[0],
                           command=lambda: [MA(), G_chantier()])
check_F1 = ttk.Checkbutton(frm_choix, text=Chantier_index[1], variable=Chantier[1],
                           command=lambda: [F1(), G_chantier()])
check_F3 = ttk.Checkbutton(frm_choix, text=Chantier_index[2], variable=Chantier[2],
                           command=lambda: [F3(), G_chantier()])

check_MA.grid(row=1, column=2)
check_F1.grid(row=2, column=2)
check_F3.grid(row=3, column=2)

# Courbes
Courbe = [tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar()]
Courbe_index = ["indice AFS", "pH", "Poussières (somme tamis 10 à 12)", "Perte au feu", "Tamis refus n° :", "Flexion",
                "Glissement"]
Courbe_legende = ["indice AFS", "pH\nAttention il s'agit d'une échelle en log10 de base",
                  "Poussières (somme tamis 10 à 12)\n en gramme", "Perte au feu\n en %", "refus des tamis\n en gramme",
                  "Contrainte Flexion en Mpa (*10)\n par exemple 25 => 2,5 Mpa", "Grammes par tamis"]


def G_ord():
    B_ord = [check_AFS, check_pH, check_pouss, check_feu, check_refus, check_Flexion, check_glissement]
    k = 0
    for i in range(len(Courbe)):
        k = k + Courbe[i].get()
    if k == 0:
        for i in range(len(Courbe)):
            B_ord[i].state(['!disabled'])
        return
    elif k == 1:
        for i in range(len(Courbe)):
            if Courbe[i].get() == 0:
                B_ord[i].state(['disabled', '!selected'])
    return


# indice AFS \n correspond au numéro d'un tamis fictif qui retiendrait\ntout le sable si les grains étaient de même dimensions moyenne
check_AFS = ttk.Checkbutton(frm_choix, text=Courbe_index[0], variable=Courbe[0], command=G_ord)
check_pH = ttk.Checkbutton(frm_choix, text=Courbe_index[1], variable=Courbe[1], command=G_ord)
check_pouss = ttk.Checkbutton(frm_choix, text=Courbe_index[2], variable=Courbe[2], command=G_ord)
check_feu = ttk.Checkbutton(frm_choix, text=Courbe_index[3], variable=Courbe[3], command=G_ord)
check_refus = ttk.Checkbutton(frm_choix, text=Courbe_index[4], variable=Courbe[4], command=G_ord)


# Fonction déacivant les boutons inutiles en parralèle d'une demande de flexion
def clear_flexion():
    if Courbe[5].get() == 1:
        k = 0
        for i in [check_sable_recup, check_sable_RTH, check_silice_neuve, check_chromite_neuve, check_chromite_recup,
                  check_silo, check_sablerie, check_malaxeur, check_BB, check_vrac]:
            i.state(['disabled', '!selected'])
            if k < 5:
                Sable[k].set(0)
                Prelev[k].set(0)
            k = k + 1
        MA()
        F1()
        F3()
    elif Courbe[5].get() == 0:
        for i in [check_sable_recup, check_sable_RTH, check_silice_neuve, check_chromite_neuve, check_chromite_recup,
                  check_silo, check_sablerie, check_malaxeur, check_BB, check_vrac]:
            i.state(['!disabled'])
        k = 0
        for i in [check_MA2, check_IMF_MA, check_FAT, check_IMF_F1, check_SAPIC, check_FAT_F3, check_FAT_F5]:
            i.state(['disabled', '!selected'])
        if k < 6:
            Flexion_malaxeur[k].set(0)
        k = k + 1
    return


check_Flexion = ttk.Checkbutton(frm_choix, text="Flexion", variable=Courbe[5],
                                command=lambda: [clear_flexion(), G_ord()])
check_glissement = ttk.Checkbutton(frm_choix, text="Glissement des tamis", variable=Courbe[6], command=G_ord)
tamis1 = ttk.Combobox(frm_choix, value=['', 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])
tamis2 = ttk.Combobox(frm_choix, value=['', 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])
tamis3 = ttk.Combobox(frm_choix, value=['', 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])

check_AFS.grid(row=1, column=3)
check_pH.grid(row=2, column=3)
check_pouss.grid(row=3, column=3)
check_feu.grid(row=4, column=3)
check_refus.grid(row=5, column=3)
tamis1.grid(row=6, column=3)
tamis2.grid(row=7, column=3)
tamis3.grid(row=8, column=3)
check_glissement.grid(row=0, column=4)
check_Flexion.grid(row=1, column=4)

# Flexion
Flexion_malaxeur = [tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar()]
Flexion_malaxeur_index = ["MA2", "IMF (MA)", "FAT", "IMF(F1)", "SAPIC", "FAT F3", "FAT F5"]
check_MA2 = ttk.Checkbutton(frm_choix, text=Flexion_malaxeur_index[0], variable=Flexion_malaxeur[0])
check_IMF_MA = ttk.Checkbutton(frm_choix, text=Flexion_malaxeur_index[1], variable=Flexion_malaxeur[1])
check_FAT = ttk.Checkbutton(frm_choix, text=Flexion_malaxeur_index[2], variable=Flexion_malaxeur[2])
check_IMF_F1 = ttk.Checkbutton(frm_choix, text=Flexion_malaxeur_index[3], variable=Flexion_malaxeur[3])
check_SAPIC = ttk.Checkbutton(frm_choix, text=Flexion_malaxeur_index[4], variable=Flexion_malaxeur[4])
check_FAT_F3 = ttk.Checkbutton(frm_choix, text=Flexion_malaxeur_index[5], variable=Flexion_malaxeur[5])
check_FAT_F5 = ttk.Checkbutton(frm_choix, text=Flexion_malaxeur_index[6], variable=Flexion_malaxeur[6])

for i in [check_MA2, check_IMF_MA, check_FAT, check_IMF_F1, check_SAPIC, check_FAT_F3, check_FAT_F5]:
    i.grid(row=[check_MA2, check_IMF_MA, check_FAT, check_IMF_F1, check_SAPIC, check_FAT_F3, check_FAT_F5].index(i) + 2,
           column=4)
    i.state(['disabled'])

# Fin Frame du choix  --------------------------------------------------------------

# Frame graphique  --------------------------------------------------------------
frm_graph = tk.Frame(master=window)
frm_graph.grid(row=1, column=0, columnspan=2, sticky="nsew")


# définition des fonctions
# pour le plot 3d (les dates sur l'axe y?)
def format_date(x, pos=None):
    return dates.num2date(x).strftime('%Y-%m-%d')  # use FuncFormatter to format dates


clear = 0
count = 0


# Fonction principale, celle qui est appelé lors de l'activation du boutton "Process!"
def Process():
    global clear
    global count
    count = count + 1
    # calcul du nombre de graph demandés
    # reset des index
    Y_index_sable = []
    Y_index_prelev = []
    Y_index_chantier = []
    Y_index_courbe = []
    Y_index_malaxeur = []

    # rapatriment des index des cases dans les listes du dessus
    for i in range(len(Sable)):
        if Sable[i].get() == 1:
            Y_index_sable.append((Sable_index[i], i))
    # print(Y_index_sable)
    for i in range(len(Prelev)):
        if Prelev[i].get() == 1:
            Y_index_prelev.append((Prelev_index[i], i))
    # print(Y_index_prelev)
    for i in range(len(Chantier)):
        if Chantier[i].get() == 1:
            Y_index_chantier.append((Chantier_index[i], i))
    # print(Y_index_chantier)
    for i in range(len(Courbe)):
        if Courbe[i].get() == 1:
            Y_index_courbe.append((Courbe_index[i], 13 + i, Courbe_legende[i]))
    # print(Y_index_courbe)
    k = (0, 0, 0)
    for i in range(len(Flexion_malaxeur)):
        if Flexion_malaxeur[i].get() == 1:
            if i < 5:
                k = (3 * i + 1, 3 * i + 2, 3 * i + 3)
            elif i >= 5:
                k = (3 * i + 4, 3 * i + 5, 3 * i + 6)
            Y_index_malaxeur.append((Flexion_malaxeur_index[i], k))
    # print(Y_index_malaxeur)

    # Test d'erreur
    """  # obsolète depuis le grisage des cases
        # Si plusieurs ordonnées en même temps
    if len(Y_index_courbe)>1:
        tk.messagebox.showinfo("Erreur", "Une seule caractéristique (colonne Ordonnées) possible")
        return
        # Si trop de variable sont sélectionées
    if (len(Y_index_sable)>1 and (len(Y_index_prelev)+len(Y_index_chantier)>2)):
        tk.messagebox.showinfo("Erreur", "Une seule des colonnes -Sable-Prelévement-Chantier- peut avoir une sélection multiple !!")
        return
    if (len(Y_index_prelev)>1 and (len(Y_index_sable)+len(Y_index_chantier)>2)):
        tk.messagebox.showinfo("Erreur", "Une seule des colonnes -Sable-Prelévement-Chantier- peut avoir une sélection multiple !!")
        return
    if (len(Y_index_chantier)>1 and (len(Y_index_prelev)+len(Y_index_sable)>2)):
        tk.messagebox.showinfo("Erreur", "Une seule des colonnes -Sable-Prelévement-Chantier- peut avoir une sélection multiple !!")
        return
    """
    # Si pas assez de variables sont sélectionées
    if len(Y_index_sable) == 0 or len(Y_index_prelev) == 0 or len(Y_index_chantier) == 0 or len(Y_index_courbe) == 0:
        if Y_index_courbe[0][1] != 18 and Y_index_courbe[0][
            1] != 19:  # si glissement ou flexion demandé on affiche pas l'erreur
            tk.messagebox.showinfo("Erreur", "Veuilliez sélectioner au moins une case dans chaque colonne")
            return
    # Si un champ de date est vide on prends 1 pour le jour, 1 pour le mois et 2020 pour l'année
    Index_date = [Jours_start, Mois_start, Jours_end, Mois_end]
    if Ans_start.get() == "":
        Ans_start.set("2020")
    if Ans_end.get() == "":
        Ans_end.set("2021")
    for i in range(len(Index_date)):
        if Index_date[i].get() == "":
            Index_date[i].set("1")

    # l'axe X
    start = dt.datetime(int(Ans_start.get()), int(Mois_start.get()[0:2]), int(Jours_start.get()))
    end = dt.datetime(int(Ans_end.get()), int(Mois_end.get()[0:2]), int(Jours_end.get()))
    # Si Erreur date pas dans l'ordre
    if start > end:
        tk.messagebox.showinfo("Erreur", "Les dates ne sont pas correctement rentrées")
        return

    X = mdates.drange(start, end, dt.timedelta(days=1))
    Xdate = mdates.num2date(X)
    # print(Xdate[0],Xdate[1])
    X_graph = []
    X_graph_flexion = [[], [], []]

    # l'axe Y

    Y = []
    Y_tamis = [[], [], []]
    Y_flexion = [[], [], []]
    tamis = []  # pour le compte de cb de tamis sont demandés dans le cas du choix des glissement de tamis
    X_3d = []
    Y_3d = []
    dZ_3d = []
    color_3d = []
    color_3d_color = ["black", "black", "black", "black", "black", "orange", "red", "green", "orange", "black", "black",
                      "black"]  # couleur des barres des 12 tamis

    if Y_index_courbe[0][1] == 17:
        # SI LE GLISSEMENT DES TAMIS EST COCHé ---------------------------------
        if tamis1.get() != '':  # si le tamis 1 est référencé
            tamis.append(int(tamis1.get()))
            if tamis2.get() != '':  # si le tamis 1 et 2 sont référencés
                tamis.append(int(tamis2.get()))
                if tamis3.get() != '':  # si le tamis 1 et 2 et 3 sont référencés
                    tamis.append(int(tamis3.get()))
            elif tamis3.get() != '':  # si le tamis 1 et 3 sont référencés
                tamis.append(int(tamis3.get()))
        elif tamis2.get() != '':  # si le 2 est alors que 1 est pas
            tamis.append(int(tamis2.get()))
            if tamis3.get() != '':  # si 3 est et 2 est alors que 1 est pas
                tamis.append(int(tamis3.get()))
        elif tamis3.get() != '':
            tamis.append(int(tamis3.get()))
        else:
            tk.messagebox.showinfo("Erreur",
                                   "Vous avez demandé d'avoir le refus de tamis spécifiques, il faut sélectioner les tamis voulus (1, 2 ou 3 tamis max)")
            return
        # FIN SI LE GLISSEMENT DES TAMIS EST COCHé ---------------------------------

    # Dans le cas qu'une seule courbe soit demandée
    if len(Y_index_sable) + len(Y_index_prelev) + len(Y_index_chantier) == 3 or Y_index_courbe[0][1] == 18:
        if Y_index_courbe[0][1] != 18 and Y_index_courbe[0][1] != 19:  # cas général
            for i in range(len(X)):
                Xdate[i] = datetime.fromordinal(Xdate[
                                                    i].toordinal())  # pour pouvoir comparer les deux dates il faut les mettre au même format (erreur avec les heures...)
                # print(datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][0][0]) - 2))
                for j in range(len(S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]])):
                    # print(datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][0]) - 2))
                    # Si le numéro de la date Exelc du docuement lu et égale à la date de Xdate[i] alors : (en gros si tu trouves un execl pour la date en question)
                    if datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(
                            S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][0]) - 2) == Xdate[
                        i]:

                        if Y_index_courbe[0][1] == 17:  # Si des tamis spécifiques sont demandés
                            # print(tamis)
                            k = 0
                            for t in range(len(tamis)):
                                if S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][
                                    tamis[t]] == 0.0:
                                    # Y_tamis[t].append(np.nan)  # sinon discontinue
                                    k = 1
                                    continue
                                Y_tamis[t].append(
                                    S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][tamis[t]])
                            if k == 0:  # sinon discontinue
                                X_graph.append(datetime.fromordinal(Xdate[
                                                                        i].toordinal()))  # pour pouvoir comparer les deux dates il faut les mettre au même format (erreur avec les heures...)
                            # print(Y_tamis)
                            break
                        if Y_index_courbe[0][1] == 13 and \
                                S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][
                                    Y_index_courbe[0][
                                        1]] == '':  # pour contrer le soucis du test > 60 après (int('') = erreur)
                            break
                        if Y_index_courbe[0][1] == 13 and (
                                S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][
                                    Y_index_courbe[0][1]] == 0.0 or
                                S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][
                                    Y_index_courbe[0][
                                        1]] > 60):  # si AFS demandé et le point non conforme (0 ou >80) alors prends pas le point
                            break
                        if Y_index_courbe[0][1] == 14 and (
                                S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][
                                    Y_index_courbe[0][1]] == '' or int(
                                S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][
                                    Y_index_courbe[0][
                                        1]]) > 14):  # si pH demandé et le point non conforme (rien ou >14) alors prends pas le point
                            break
                        if Y_index_courbe[0][1] == 16 and (
                                S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][
                                    Y_index_courbe[0][1]] == '' or int(
                                S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][
                                    Y_index_courbe[0][
                                        1]]) > 15):  # si la perte au feu demandé et le point non conforme (rien ou >15%) alors prends pas le point
                            break
                        if Y_index_courbe[0][1] == 15 and int(
                                S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][
                                    Y_index_courbe[0][1]] == '') > 5:  # si poussière > 5g
                            break
                        if Y_index_courbe[0][1] == 16:
                            Y.append(S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][
                                         Y_index_courbe[0][1]] * 100)  # pour avoir des valeurs en %
                            X_graph.append(datetime.fromordinal(Xdate[i].toordinal()))
                            continue
                        # on enregistre le point pour le plot CAS GENERAL
                        Y.append(S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][
                                     Y_index_courbe[0][1]])
                        X_graph.append(datetime.fromordinal(Xdate[
                                                                i].toordinal()))  # pour pouvoir comparer les deux dates il faut les mettre au même format (erreur avec les heures...)
                        break
        elif Y_index_courbe[0][1] == 18:  # Si les flexions sont demandés
            for i in range(len(X)):
                Xdate[i] = datetime.fromordinal(Xdate[i].toordinal())
                for j in range(len(SF)):
                    # print(Xdate[i], datetime.fromordinal(SF[j][0].toordinal()))
                    if datetime.fromordinal(SF[j][0].toordinal()) == Xdate[i]:
                        if str(SF[j][Y_index_malaxeur[0][1][0]]) != "nan":
                            Y_flexion[0].append(SF[j][Y_index_malaxeur[0][1][0]])
                            X_graph_flexion[0].append(datetime.fromordinal(Xdate[i].toordinal()))
                        if str(SF[j][Y_index_malaxeur[0][1][1]]) != "nan":
                            Y_flexion[1].append(SF[j][Y_index_malaxeur[0][1][1]])
                            X_graph_flexion[1].append(datetime.fromordinal(Xdate[i].toordinal()))
                        if str(SF[j][Y_index_malaxeur[0][1][2]]) != "nan":  # pour discontinue
                            Y_flexion[2].append(SF[j][Y_index_malaxeur[0][1][2]])
                            X_graph_flexion[2].append(datetime.fromordinal(Xdate[i].toordinal()))

        ##                        Y_flexion[0].append(SF[j][Y_index_malaxeur[0][1][0]])
        ##                        Y_flexion[1].append(SF[j][Y_index_malaxeur[0][1][1]])
        ##                        Y_flexion[2].append(SF[j][Y_index_malaxeur[0][1][2]])
        ##                        X_graph.append(datetime.fromordinal(Xdate[i].toordinal()))
        # print(Y_flexion)
        elif Y_index_courbe[0][1] == 19:  # si glissement est demandé
            for i in range(len(X)):
                Xdate[i] = datetime.fromordinal(Xdate[i].toordinal())  # Z
                nb_tamis = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]  # le numéro des tamis
                for j in range(len(S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]])):
                    if datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(
                            S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][0]) - 2) == Xdate[
                        i]:
                        for k in nb_tamis:
                            color_3d.append(color_3d_color[k - 1])
                            X_3d.append(k)
                            Y_3d.append(datetime.fromordinal(Xdate[i].toordinal()))
                            dZ_3d.append(S[Y_index_sable[0][1]][Y_index_prelev[0][1]][Y_index_chantier[0][1]][j][k])

    # print(" {0} \n {1} \n {2}".format(X_3d, Y_3d, dZ_3d))
    # print(Y_flexion)
    # print(Y)
    if len(Y) == 0 and Y_index_courbe[0][1] != 17 and Y_index_courbe[0][1] != 18 and Y_index_courbe[0][1] != 19:
        tk.messagebox.showinfo("Erreur",
                               "Il n'y a pas de données pour :\n {0} entre {1} / {2} et {3} / {4} \n {5} au {6} à {7}".format(
                                   Y_index_courbe[0][0], Mois_start.get()[0:2], Ans_start.get(), Mois_end.get()[0:2],
                                   Ans_end.get(), Y_index_sable[0][0], Y_index_prelev[0][0], Y_index_chantier[0][0]))
        return

    # PLOT DU GRAPH

    if Y_index_courbe[0][1] == 17:  # si glissement tamis demandé
        plt.figure(count)
        color = ["r", "g", "b"]
        for i in range(len(tamis)):
            # print(tamis)
            # print(Y_tamis)
            plt.plot_date(X_graph, Y_tamis[i], "+:", linewidth=1, color=color[i], xdate=True,
                          label="{0} au {1} à {2} - TAMIS N°{3}".format(Y_index_sable[0][0], Y_index_prelev[0][0],
                                                                        Y_index_chantier[0][0], tamis[i]))
            # regression linéaire
            X_trend = mdates.date2num(X_graph)
            z = np.polyfit(X_trend, Y_tamis[i], 10)  # 3 eme argument : degree de l'approximation polynomiale
            p = np.poly1d(z)
            zm = np.polyfit(X_trend, Y_tamis[i], 0)
            pm = np.poly1d(zm)
            plt.plot_date(X_trend, p(X_trend), "r-", linewidth=2, color=color[i])
            plt.plot_date(X_trend, pm(X_trend), "g-.", linewidth=1, color=color[i])
            plt.text(X_trend[0], pm(X_trend)[0], f'{pm(X_trend)[0]:.2f}', fontsize=16)
        plt.ylabel(Y_index_courbe[0][2])
        plt.title("Evolution {0} entre {1} / {2} et {3} / {4}".format(Y_index_courbe[0][0], Mois_start.get()[0:2],
                                                                      Ans_start.get(), Mois_end.get()[0:2],
                                                                      Ans_end.get()))
        plt.legend(loc=0)
        plt.gcf().autofmt_xdate()
        plt.grid()
        plt.show()
    elif Y_index_courbe[0][1] == 18:  # Si Flexion demandé
        plt.figure(count)
        Flexion_legende = ['Silice neuve', '60/40', 'Chromite neuve', 'Silice récup 25T', 'Silice Récup 50T', 'RTH',
                           'Silice récup', '60/40', '', 'Silice récup 25T', 'Silice récup 50T', '', 'Silice neuve',
                           'Silice récup', '', '', '', '', 'Silice neuve', 'Silice récup', 'Chromite récup', '',
                           'Silice récup', '', '', '', '']
        color = ["g", "r", "y"]
        for i in range(len(Y_flexion)):
            if len(Y_flexion[i]) == 0:  # si la liste à afficher est  vide
                continue
            plt.plot_date(X_graph_flexion[i], Y_flexion[i], "+:", linewidth=1, xdate=True, color=color[i],
                          label="- {0}".format(Flexion_legende[Y_index_malaxeur[0][1][i] - 1]))
            # regression linéaire
            X_trend = mdates.date2num(X_graph_flexion[i])
            z = np.polyfit(X_trend, Y_flexion[i], 10)  # 3 eme argument : degree de l'approximation polynomiale
            p = np.poly1d(z)
            zm = np.polyfit(X_trend, Y_flexion[i], 0)
            pm = np.poly1d(zm)
            plt.plot_date(X_trend, p(X_trend), "r-", linewidth=2, color=color[i])
            plt.plot_date(X_trend, pm(X_trend), "g-.", linewidth=1, color=color[i])
            plt.text(X_trend[0], pm(X_trend)[0], f'{pm(X_trend)[0]:.2f}', fontsize=16)
        plt.ylabel(Y_index_courbe[0][2])
        plt.title(
            "Evolution {0} entre {1} / {2} et {3} / {4} sur le chantier {5} malaxeur {6}".format(Y_index_courbe[0][0],
                                                                                                 Mois_start.get()[0:2],
                                                                                                 Ans_start.get(),
                                                                                                 Mois_end.get()[0:2],
                                                                                                 Ans_end.get(),
                                                                                                 Y_index_chantier[0][0],
                                                                                                 Y_index_malaxeur[0][
                                                                                                     0]))
        plt.legend(loc=0)
        plt.gcf().autofmt_xdate()
        plt.grid()
        plt.show()
    elif Y_index_courbe[0][1] == 19:  # Si glissement 3d demandé
        fig = plt.figure(figsize=(5, 5), dpi=100)
        # print(X_3d,"\n",mdates.date2num(Y_3d),"\n",np.zeros_like(dZ_3d),"\n",dZ_3d)
        fig.add_subplot(111, projection='3d').bar3d(X_3d, mdates.date2num(Y_3d), np.zeros_like(dZ_3d), 0.5, 2, dZ_3d,
                                                    shade=True, color=color_3d)
        fig.gca(projection="3d").yaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        plt.title(
            "Glissement granulométrique entre {0} / {1} et {2} / {3}".format(Mois_start.get()[0:2], Ans_start.get(),
                                                                             Mois_end.get()[0:2], Ans_end.get()))
        plt.xlabel("Numéro tamis")
        fig.gca(projection='3d').set_zlabel(Y_index_courbe[0][2])
        plt.show()
    else:
        plt.figure(count)
        if Sable[2].get() == 1 or Sable[3].get() == 1:
            plt.plot_date(X_graph, Y, "+:", linewidth=1, xdate=True, label="{0}".format(Y_index_sable[0][0]))
        else:
            plt.plot_date(X_graph, Y, "+:", linewidth=1, xdate=True,
                          label="{0} au {1} à {2}".format(Y_index_sable[0][0], Y_index_prelev[0][0],
                                                          Y_index_chantier[0][0]))
        # regression linéaire
        X_trend = mdates.date2num(X_graph)
        z = np.polyfit(X_trend, Y, 10)  # 3 eme argument : degree de l'approximation polynomiale
        p = np.poly1d(z)
        zm = np.polyfit(X_trend, Y, 0)
        pm = np.poly1d(zm)
        plt.plot_date(X_trend, p(X_trend), "-", linewidth=2)  # , label = "Aproximation polynomiale d'ordre 10")
        plt.plot_date(X_trend, pm(X_trend), "-.", linewidth=1)  # , label = "Moyenne")
        plt.text(X_trend[0], pm(X_trend)[0], f'{pm(X_trend)[0]:.2f}', fontsize=16)
        plt.ylabel(Y_index_courbe[0][2])
        plt.title("Evolution {0} entre {1} / {2} et {3} / {4}".format(Y_index_courbe[0][0], Mois_start.get()[0:2],
                                                                      Ans_start.get(), Mois_end.get()[0:2],
                                                                      Ans_end.get()))
        plt.legend(loc=0)
        plt.gcf().autofmt_xdate()
        plt.grid()
        plt.show()


bt_graph = tk.Button(master=frm_graph, text="Process !", command=Process).pack()
lb_contact = tk.Label(master=frm_graph, text=" contactez moi pour toute amélioration ou problème : j.thivet@esff.fr",
                      font=("Calibri", 10)).pack()

window.mainloop()
