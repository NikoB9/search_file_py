#lancer la commande suivante dans un terminal pour générer le txt des fichiers à copier grace au script python
#find /media/lubuntu -name '*.doc' -o -name '*.docx' -o -name '*.pdf' -o -name '*.JPG' -o -name '*.JPEG' -o -name '*.jpg' -o -name '*.PNG' -o -name '*.jpeg' -o -name '*.png' >> /home/lubuntu/Desktop/cp.txt

import os
import os.path
import shutil

#on ouvre le fichier généré précédemment
fichier = open("/home/lubuntu/Desktop/cp.txt","r")

#On fait défiler les lignes
for ligne in fichier:
    
    #on retire le "\n" en fin de ligne
    ligneSplit=ligne[:-1]
    print(ligneSplit)

    #on récupere le nom de fichier et son chemin
    cpfichier=os.path.basename(ligneSplit)
    cpdir=os.path.dirname(ligneSplit)

    #on essaie de créer le chemin du fichier sur la clé usb "safeKey" s'il n'existe pas
    try :
        os.makedirs("/media/lubuntu/safeKey"+cpdir)
    except OSError as err :
        print("OS error : {0}".format(err))
    
    #on copie le fichier de la machine vers la clé usb
    try:
        shutil.copy(ligneSplit, "/media/lubuntu/safeKey"+ligneSplit)
    except :
	pass