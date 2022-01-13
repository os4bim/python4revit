# -*- coding: utf-8 -*-

__doc__ = "Création automatique d'un bordereau de prix pour gaines de ventilation\nPrérequis :\n- Ouvrir une feuille Excel\n- Avoir au moins un segment, un fitting et un Accessory par cicuit"
__title__ = 'Export\nDucts'
__author__ = 'Yoann OBRY'

#BOM to Excel Ducts v1.0
#Company CERN

import clr
import System
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import * 

from System import Guid


app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document

import math

#Shared parameter code circuit
code_cir = Guid(r'55934d0c-0246-4ce2-9bdf-57ed4244e11b')

#Shared parameter FMF_Angle
angle = Guid(r'a8b84336-4f16-462c-a50f-f0f8b2e4f7c2')

### DA : Création d'un BOM de DUCT ACCESSORIES sous forme de liste de tuple

#Collecte les Duct Accessories
DAs = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()

#Créer des listes vides
DA_code_circuit = []
DA_family_name = []
DA_description = []
DA_size = []

for DA in DAs:

	
	## Get Type Parameter value
	DA_type = doc.GetElement(DA.GetTypeId())
	
	# Element ID - Instance Parameter
	#print DA.Id

	# Code circuit - Instance Parameter (Shared Parameter)
	code_circuit = DA.get_Parameter(code_cir)
	DA_code_circuit.append(code_circuit.AsString())

	# Family Name - Type Parameter
	family_name = DA_type.get_Parameter(
					BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
	DA_family_name.append(family_name.AsString())

	# Description - Type Parameter
	description = DA_type.get_Parameter(
					BuiltInParameter.ALL_MODEL_DESCRIPTION)
	DA_description.append(description.AsString())
	
	# Size - Instance Parameter
	size = DA.get_Parameter(
					BuiltInParameter.RBS_CALCULATED_SIZE)
	DA_size.append(size.AsString())

## Change les valeurs 'None' et '' en 'N/A'
for i in range(len(DA_code_circuit)):
    if DA_code_circuit[i] == None or DA_code_circuit[i] == '':
        DA_code_circuit[i] = '_N/A'

## Assemblage des listes de caractéristiques en une seule
DA_libelle = [DA_family_name[i] +"  "+ DA_description[i] +"  "+ DA_size[i] for i in range(len(DA_code_circuit))]

## Identification des codes circuits
circuit_unique = set(DA_code_circuit)
circuit_unique = list(circuit_unique)

## Créer une liste par élément avec unité de mesure et count=1
lstDA = [[DA_code_circuit[i],DA_libelle[i],'u',1] for i in range(len(DA_code_circuit))]

## Compte le nombre d'éléments identique
DAcount=[]
for i in range(len(lstDA)):
    DAcount.append(lstDA.count(lstDA[i]))
## Incrémente les quantité tout en conservant les doublons
for i in range(len(lstDA)):
    lstDA[i][3]=DAcount[i]
    
## Supprime les doublons
setDA=set(tuple(row) for row in lstDA)
lstDA=list(setDA)
lstDA.sort()

if not lstDA:
	lstDA.append("Nulle")

print(lstDA)


### DT : Création d'un BOM de DUCT SEGMENTS sous forme de liste de tuple

#Collecte les Ducts
DTs = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()

#Créer des listes vides
DT_code_circuit = []
DT_type_name = []
DT_size = []
DT_length = []

for DT in DTs:

	
	## Get Type Parameter value
	DT_type = doc.GetElement(DT.GetTypeId())
	
	# Element ID - Instance Parameter
	#print DT.Id

	# Code circuit - Instance Parameter (Shared Parameter)
	code_circuit = DT.get_Parameter(code_cir)
	DT_code_circuit.append(code_circuit.AsString())

	# Type Name - Type Parameter
	type_name = DT_type.get_Parameter(
					BuiltInParameter.SYMBOL_NAME_PARAM)
	DT_type_name.append(type_name.AsString())

	# Size - Instance Parameter
	size = DT.get_Parameter(
					BuiltInParameter.RBS_CALCULATED_SIZE)
	DT_size.append(size.AsString())

	# Length - Instance Parameter
	length = DT.get_Parameter(
					BuiltInParameter.CURVE_ELEM_LENGTH)
	DT_length.append(length.AsDouble())

## Change les valeurs 'None' et '' en 'N/A'
for i in range(len(DT_code_circuit)):
    if DT_code_circuit[i] == None or DT_code_circuit[i] == '':
        DT_code_circuit[i] = '_N/A'


## Assemblage des listes de caractéristiques en une seule
DT_libelle = [DT_type_name[i] +"  "+ DT_size[i] for i in range(len(DT_code_circuit))]


## Créer une liste par élément avec unité de mesure et métré total
lstDT = [[DT_code_circuit[i],DT_libelle[i],DT_length[i]/3.2808] for i in range(len(DT_code_circuit))]

lstDT_unique = list(set([(element[0],element[1]) for element in lstDT]))

quantites = [sum([float(part[2]) for part in lstDT if (part[0],part[1]) == element]) for element in lstDT_unique]

lstDT = [list(lstDT_unique[element])+["{:01.1f}".format(quantites[element])] for element in range(0,len(lstDT_unique))]
lstDT = [[lstDT[i][0],lstDT[i][1],'m',lstDT[i][2]] for i in range(len(lstDT))]

lstDT.sort()

print(lstDT)


### DF : Création d'un BOM de DUCT FITTINGS sous forme de liste de tuple

#Collecte les Ducts Fittings
DFs = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()

#Créer des listes vides
DF_code_circuit = []
DF_family_name = []
DF_type_name = []
DF_size = []
DF_angle = []

for DF in DFs:

	
	## Get Type Parameter value
	DF_type = doc.GetElement(DF.GetTypeId())
	
	# Element ID - Instance Parameter
	#print PF.Id

	# Code circuit - Instance Parameter (Shared Parameter)
	code_circuit = DF.get_Parameter(code_cir)
	DF_code_circuit.append(code_circuit.AsString())

	# Family Name - Type Parameter
	family_name = DF_type.get_Parameter(
					BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
	DF_family_name.append(family_name.AsString())

	# Type Name - Type Parameter
	type_name = DF_type.get_Parameter(
					BuiltInParameter.SYMBOL_NAME_PARAM)
	DF_type_name.append(type_name.AsString())
	
	# Size - Instance Parameter
	size = DF.get_Parameter(
					BuiltInParameter.RBS_CALCULATED_SIZE)
	DF_size.append(size.AsString())

	# Angle	- Instance Parameter (Shared Parameter)
	angle_coude = DF.get_Parameter(angle)
	if angle_coude:
		DF_angle.append(angle_coude.AsDouble() * 180 / math.pi)
	else:
		DF_angle.append(0)

## Arrondi les angles des ducts fittings		
for i in range(len(DF_angle)):
    if 85 <= DF_angle[i] <= 95:
        DF_angle[i] = 90

for i in range(len(DF_angle)):
    if 55 <= DF_angle[i] <= 65:
        DF_angle[i] = 60	
		
for i in range(len(DF_angle)):
    if 40 <= DF_angle[i] <= 50:
        DF_angle[i] = 45
		
for i in range(len(DF_angle)):
    if 25 < DF_angle[i] <= 35:
        DF_angle[i] = 30
		
for i in range(len(DF_angle)):
    if 15 <= DF_angle[i] <= 25:
        DF_angle[i] = 20		

## Change les valeurs 'None' et '' en 'N/A'
for i in range(len(DF_code_circuit)):
    if DF_code_circuit[i] == None or DF_code_circuit[i] == '':
        DF_code_circuit[i] = '_N/A'



## Assemblage des listes de caractéristiques en une seule
DF_libelle = [DF_family_name[i] +"  "+ DF_type_name[i] +"  "+ DF_size[i] +"  "+ str("{:01.0f}".format(5 * round(DF_angle[i])/5)) +"°" for i in range(len(DF_code_circuit))]

## Efface les angles nuls dans le libellé
DF_libelle = [w.replace('  0°','') for w in DF_libelle]


## Identification des codes circuits
circuit_unique = set(DF_code_circuit)
circuit_unique = list(circuit_unique)
circuit_unique.sort()

## Créer une liste DF d'éléments avec unité de mesure et count=1
lstDF = [[DF_code_circuit[i],DF_libelle[i],'u',1] for i in range(len(DF_code_circuit))]

## Compte le nombre d'éléments identique
DFcount=[]
for i in range(len(lstDF)):
    DFcount.append(lstDF.count(lstDF[i]))
## Incrémente les quantité tout en conservant les doublons
for i in range(len(lstDF)):
    lstDF[i][3]=DFcount[i]
    
## Supprime les doublons
setDF=set(tuple(row) for row in lstDF)
lstDF=list(setDF)
lstDF.sort()

print(lstDF)



		### Exporter les données dans Excel ###

#Command write in excel
t = Transaction(doc, 'Write Excel.')
 
t.Start()
 
#Accessing the Excel applications.
xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject('Excel.Application')
 
#Worksheet, Row, and Column parameters
worksheet = 1
rowStart = 1
columnStart = 1
 
#Effacer la feuille excel
for i in range(100):
	for j in range(10):
		data = xlApp.Worksheets(worksheet).Cells(rowStart + i, columnStart + j)
		data.Value = ""
 
#Compteur de lignes excel
count_circuit = 0
saut_ligne = 0

#Fonction qui permet à i de commencer à 0 pour l'écriture des circuits suivants
def find(c,d):
	return [(i, premier.index(c)) for i, premier in enumerate(d) if c in premier]

##Exceptions de l'Index Error

for i in range(len(circuit_unique)):
	try:
		lstDA[i][0]
	except IndexError:
		print("Chaque circuit doit contenir au moins un Duct, un Duct fitting et un Duct Accessory") #Valeur attendu : 'R03'
		
		
for i in range(len(circuit_unique)):
	try:
		find(circuit_unique[i],lstDA)[0][0]
	except IndexError:
		print("Chaque circuit doit contenir au moins un Duct, un Duct fitting et un Duct Accessory")	#Valeur attendu : 'Numéro de l'index dans lstDA pour le premier DA 'R03'
	

for k in range(len(circuit_unique)):

	count_lstDA = 0
	count_lstDT = 0
	count_lstDF = 0

	## Numéro Circuit
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + saut_ligne, columnStart)
	data.Value = range(len(circuit_unique))[k] + 1
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + saut_ligne, columnStart + 1)
	data.Value = "Circuit - " + circuit_unique[k]


	## Ecriture des Duct Accessories

	# Titre
	saut_ligne += 2
	
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + saut_ligne, columnStart)
	data.Value = range(len(circuit_unique))[k] + 1 + 0.1
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + saut_ligne, columnStart + 1)
	data.Value = "Accessoires de gaines et instrumentation"

	# Eléments
	saut_ligne += 1
	decal = find(circuit_unique[k],lstDA)[0][0]
	for i in range(len(lstDA)):

		if lstDA[i][0] == circuit_unique[k]:
			#Worksheet object specifying the cell location.
			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + i - decal + saut_ligne, columnStart + 6)
			#Assigning a value to the cell.
			data.Value = lstDA[i][0]
		
			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + i - decal + saut_ligne, columnStart + 1)
			data.Value = lstDA[i][1]

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + i - decal + saut_ligne, columnStart + 2)
			data.Value = lstDA[i][2]

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + i - decal + saut_ligne, columnStart + 3)
			data.Value = lstDA[i][3]

			count_lstDA += 1



	## Ecriture des Ducts

	# Titre
	saut_ligne += 1
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + saut_ligne, columnStart)
	data.Value = range(len(circuit_unique))[k] + 1 + 0.2
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + saut_ligne, columnStart + 1)
	data.Value = "Longueurs de gaines"

	#Eléments
	saut_ligne += 1
	decal = find(circuit_unique[k],lstDT)[0][0]
	for i in range(len(lstDT)):

		if lstDT[i][0] == circuit_unique[k]:

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + saut_ligne + i - decal, columnStart + 6)
			data.Value = lstDT[i][0]

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + saut_ligne + i - decal, columnStart + 1)
			data.Value = lstDT[i][1]
		 
			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + saut_ligne + i - decal, columnStart + 2)
			data.Value = lstDT[i][2]
		 
			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + saut_ligne + i - decal, columnStart + 3)
			data.Value = lstDT[i][3]
			
			count_lstDT += 1


	## Ecriture des Duct Fittings
	 
	# Titre
	saut_ligne += 1
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + count_lstDT + saut_ligne, columnStart)
	data.Value = range(len(circuit_unique))[k] + 1 + 0.3
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + count_lstDT + saut_ligne, columnStart + 1)
	data.Value = "Raccords"

	#Eléments
	saut_ligne += 1
	decal = find(circuit_unique[k],lstDF)[0][0]
	for i in range(len(lstDF)):

		if lstDF[i][0] == circuit_unique[k]:

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + count_lstDT + saut_ligne + i - decal, columnStart + 6)
			data.Value = lstDF[i][0]

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + count_lstDT + saut_ligne + i - decal, columnStart + 1)
			data.Value = lstDF[i][1]

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + count_lstDT + saut_ligne + i - decal, columnStart + 2)
			data.Value = lstDF[i][2]
		 
			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + count_lstDT + saut_ligne + i - decal, columnStart + 3)
			data.Value = lstDF[i][3]
			
			count_lstDF += 1

	## Sous total
	saut_ligne += 1
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + count_lstDT + count_lstDF + saut_ligne, columnStart)
	data.Value = "ST" + str(range(len(circuit_unique))[k] + 1)
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstDA + count_lstDT + count_lstDF + saut_ligne, columnStart + 1)
	data.Value = "Total " + str(range(len(circuit_unique))[k] + 1) + " sous poste"

	count_circuit += count_lstDA + count_lstDT + count_lstDF
	saut_ligne += 2
	
	
	
t.Commit()
