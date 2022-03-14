# Python scripts for Revit

🗣️ FR

Ces fichiers sont utilisé via pyrevit. Pour les utiliser via RevitPythonShell supprimer ces lignes :
```
__doc__ =
__title__ =
__author__ =
```

### BOM_to_Excel_Ducts_script.py et BOM_to_Excel_Pipes_script.py

Ces scripts permettent d'extraire un quantitatif de gaines ou de tuyauterie depuis Revit pour l'écrire dans un fichier Excel, sans passer par les nomenclatures.
Cela permet de créer rapidement un bordereau de prix pour envoyer à un fournisseur.

Avant de lancer le script il faut ouvrir un fichier excel dans lequel le quantitatif sera écrit.
Le programme peut-êre améliorer car pour l'instant il ne marche pas si il n'y a pas au moins une longueur droite, un raccord et un accessoire par circuit, ce qui sera le cas dans la plupart des installations, sauf peut-être pour les évacuations.
Nous pourrions imaginer introduire un champs N/A pour ce cas.
Si vous avez une idée n'hésiter pas à me la proposer !

### Show_Parameters.py

Ce script vous permet de récupérer les noms des paramètres du premier objet Revit présent dans votre fichier. Pour ce faire vous pouvez ne laisser qu'un seul objet dans votre projet Revit pour récupérer ces données.
A noter la différence entre les éléments "type" et les éléments "instance" ou "non type". Ceci est important car il faut ajouter une ligne de code pour que le script récupère les données des élements "type".

Le fichier excel PA_Parameters_name.xlsx est un exemple de noms de paramètres Revit pour un "Pipe Accessories"

🗣️ EN

This files are used in pyrevit. For using in RevitPythonShell just delete this lines :
```
__doc__ =
__title__ =
__author__ =
```

### BOM_to_Excel_Ducts_script.py et BOM_to_Excel_Pipes_script.py

These scripts allow you to extract a duct or piping quantity from Revit and write it to an Excel file, without going through the schedules.
This allows you to quickly create a price list to send to a supplier.

Before running the script, an excel file must be opened in which the quantity will be written.
The program can be improved because at the moment it does not work if there is not at least one straight length, one fitting and one accessory per circuit, which will be the case in most installations, except perhaps for drains.
We could imagine introducing a N/A field for this case.
If you have an idea don't hesitate to suggest it to me !
