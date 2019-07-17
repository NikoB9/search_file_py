#!/usr/bin/python
# -*- coding: utf-8 -*-
#Nicolas BOURNEUF 
#16/07/2019

#librairie exploitation excel
import xlrd

#importation des utilitaire pour créer un excel
from xlwt import Workbook, Formula

#importation outils de recherche
import os
from pathlib import Path

# On créer un "classeur" pour enregistrer les résultats de la recherche
classeur = Workbook()
# On ajoute une feuille au classeur
feuilleResults = classeur.add_sheet("chemins old_CARTADS")
feuilleResults.write(0, 0, "Chemin de recherche")
feuilleResults.write(0, 1, "Chemin de resultat")

#importation du document sue lequel faire la recherche
document = xlrd.open_workbook("retrouver_fichier.xlsx")

#Nombre et nom des feuille du fichier
print("Nombre de feuilles: "+str(document.nsheets))
print("Noms des feuilles: "+str(document.sheet_names()))

#Récupération première feuille par son index
feuille = document.sheet_by_index(0) 
#PAR NOM : document.sheet_by_name("retrouver_fichier")

print("Info générales de la feuille en cours de traitement : ")
print("Nom: "+str(feuille.name))
print("Nombre de lignes: "+str(feuille.nrows))
print("Nombre de colonnes: "+str(feuille.ncols))

#écupération nombre de colonnes et de lignes 
cols = feuille.ncols
rows = feuille.nrows

#Je sais que la colonne 2 est la colonne qui comporte les chemins dont j'ai besoin
#Je vais itérer sur ma 2ème colonne (de la 2e à la dernière ligne) et stocker tous les chemins dans un tableau
pathNewCartads = []
pathOldCartads = []
for row in range(1, rows):

    #valeur / chemin du fichier à trouver
    valueLine = feuille.cell_value(rowx=row, colx=1) 
    #sauvegarde du chemin de fichier à trouver
    pathNewCartads.append(valueLine)

    #print("Ligne " + str(row) + " : " + str(chemins[row-1]))

    #récupération des information nécessaires sur le chemin récupéré
    cutPath = valueLine.split('/')
    arrondissement = cutPath[2]
    dateDos = cutPath[3]
    typeDos = cutPath[4]
    nameDos = cutPath[5]
    fileToFind = cutPath[7]

    pathToSearch = '//10.161.17.244/Documentation/'+str(arrondissement)+'/'+str(typeDos)+'/'+str(dateDos)+'/'+str(nameDos)
    #print(pathToSearch + " : ")
    #entries = os.listdir(pathToSearch)
    #for entry in entries:
    #    print(entry)
    #    if(fileToFind == entry):
    #        break
    fichierTrouve = False
    cheminFichierTrouve = "Le fichier n'a pas été trouvé"

    #Si le chemin existe on va chercher dedans
    if(os.path.exists(pathToSearch)):

        basepath = Path(pathToSearch)
        files_in_basepath = basepath.iterdir()
        #pour chaque parcours disponible
        for item in files_in_basepath:
            #si c'est un fichier on regarde si le fichier recherché est présent
            if item.is_file():
                #si on trouve le ficher on dit qu'on l'a trouvé et on l'enregistre son chemin dans une variable
                if(fileToFind == item.name):
                    cheminFichierTrouve = pathToSearch+'/'+str(item.name)
                    fichierTrouve = True
                    break
            #sinon c'est surment un dossier (.is_dir()) on relance la recherche de fichier dessus
            else :
                subpath = Path(pathToSearch+'/'+str(item.name))
                files_in_subpath = subpath.iterdir()
                for subitem in files_in_subpath:
                    if subitem.is_file():
                        #si on trouve le ficher on dit qu'on l'a trouvé et on l'enregistre son chemin dans une variable
                        if(fileToFind == subitem.name):
                            cheminFichierTrouve = pathToSearch+'/'+str(item.name)+'/'+str(subitem.name)
                            fichierTrouve = True
                            break
                    #sinon c'est surment un dossier (.is_dir()) on relance la recherche de fichier dessus
                    else :
                        subsubpath = Path(pathToSearch+'/'+str(item.name)+'/'+str(subitem.name))
                        files_in_subsubpath = subsubpath.iterdir()
                        for subsubitem in files_in_subsubpath:
                            #si on trouve le ficher on dit qu'on l'a trouvé et on l'enregistre son chemin dans une variable
                            if(fileToFind == subsubitem.name):
                                cheminFichierTrouve = pathToSearch+'/'+str(item.name)+'/'+str(subitem.name)+'/'+str(subsubitem.name)
                                fichierTrouve = True
                                break
                    

    #formater le chemin pour qu'il soit cliquable
    OLD_pathToFormate = cheminFichierTrouve.split('/')
    OLD_formatePath = ""
    for f in range(len(OLD_pathToFormate)) : 
        if (f == len(OLD_pathToFormate)-1):
            OLD_formatePath += OLD_pathToFormate[f]
        else:
            OLD_formatePath += OLD_pathToFormate[f] + "\\"

    
    # Ecrire base la colonne 0
    feuilleResults.write(row, 0, valueLine)
    # Ecrire resultat dans colonne 1
    feuilleResults.write(row, 1, OLD_formatePath)
    # Ecrire une formule dans la cellule à la ligne 0 et la colonne 2
    # qui va additioner les deux autres cellules
    #feuille.write(0, 2, Formula('A1+B1'))
    
# Ecriture du classeur sur le disque
path = r"C:\Users\bourneun\Desktop\TraitementPython\resultatRecherche_v3.xls"
classeur.save(path)
print("fin :)")
          








 

