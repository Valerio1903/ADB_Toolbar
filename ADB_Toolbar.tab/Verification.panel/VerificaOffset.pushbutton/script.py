# -*- coding: utf-8 -*-
""" Verifica la corretta collocazione altimetrica delle famiglie caricabili rispetto al livello di host  """
__author__= 'Roberto Dolfini'
__title__ = 'Verifica Offset\nElementi'

import codecs
import re
import unicodedata
import pyrevit
from pyrevit import forms, script
import math
import clr
import os
import codecs
import sys
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import csv
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *
#PER IL FLOATING POINT
from decimal import Decimal, ROUND_DOWN, getcontext

##############################################################
doc   = __revit__.ActiveUIDocument.Document  #type: Document
uidoc = __revit__.ActiveUIDocument						   
app   = __revit__.Application		
aview = doc.ActiveView
output = pyrevit.output.get_output()

t = Transaction(doc, "TROUBLESHOOTING")

##############################################################

# CREAZIONE LISTE DI OUTPUT DATA
OFFSET_CSV_DATA = []
OFFSET_CSV_DATA.append(["Categoria","ID Elemento","Verifica","Offset Rilevato","Stato"])


def ConvertiUnita(valore):
	Valore_Decimale = Decimal(UnitUtils.Convert(valore,UnitTypeId.Feet,UnitTypeId.Meters))
	getcontext().prec = 4
	#Valore_Raffinato = Valore_Decimale.quantize(Precisione_Decimale, rounding=ROUND_DOWN)
	return round(Valore_Decimale,4)

def VerificaRange(offset, range):
	if offset > range:
		return "Troppo Alto"
	elif 0 <= offset < range:
		return "Ok"
	else:
		return "Troppo Basso"
def EstraiFamigliaOggetto(oggetto):
	try:
		return oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
	except:
		try:
			return oggetto.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsValueString()
		except:
			return "NON TROVATO"

######################################################################################


output.print_md("# Verifica Offset Elementi")
output.print_md("---")

# pyRevit Script - Verifica Livelli e Posizione Caricabili

from Autodesk.Revit.DB import *
from pyrevit import script

output = script.get_output()

# Funzione: conversione unità Revit -> metri
def converti_unita(valore):
	return round(valore * 0.3048, 2)

# Funzione: calcolo altezza di un FamilyInstance
def calcola_altezza(elemento):
	CategorieDavanzali = [BuiltInCategory.OST_Windows, BuiltInCategory.OST_Doors]

	if elemento.Category.BuiltInCategory in CategorieDavanzali:
		offset = elemento.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM)

	else:
		offset = elemento.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)

	return converti_unita(offset.AsDouble())

# Funzione: recupera livelli ordinati per elevazione e suddivisi per disciplina
def recupera_livelli(doc):
	livelli = FilteredElementCollector(doc).OfClass(Level).ToElements()
	livelli_con_altitudine = [(liv.Id, liv.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble()) for liv in livelli]
	livelli_ordinati = sorted(livelli_con_altitudine, key=lambda x: x[1])

	livelli_arc, livelli_str, livelli_ignoti = [], [], []
		
	for livello_id, _ in livelli_ordinati:
		livello = doc.GetElement(livello_id)
		codice = livello.Name.split("_")[0]
		if "AR" in codice:
			livelli_arc.append(livello)
		elif "ST" in codice:
			livelli_str.append(livello)
		else:
			livelli_ignoti.append(livello)

	return livelli_arc, livelli_str, livelli_ignoti

def verifica_elevazione_superiore(elevazione, livelli, id_livello):
	if elevazione < 0:
		return 0,":cross_mark:","Offset Negativo."
		
	try:
		Interpiano = livelli[id_livello + 1].Elevation
		if elevazione > Interpiano:
			return 0,":cross_mark:","Oggetto sopra l'interpiano."
		else:
			return 1,":white_heavy_check_mark:","Oggetto conforme."
		
	except:
		if elevazione > 15: #! valore forfettario aggiunto per evitare modellazioni a caso
			return 0,":cross_mark:","Oggetto con elevazione troppo alta."
		else:
			return 1,":white_heavy_check_mark:","Oggetto conforme."


# Funzione: stampa livelli per disciplina
def stampa_livelli(label, lista):
	output.print_md("***Livelli {}:***".format(label))
	for livello in lista:
		output.print_md("**{}** : {}m".format(livello.Name, converti_unita(livello.Elevation)))
	output.print_md("---")

def recupera_elementi(doc):
	categorie_escluse = [BuiltInCategory.OST_CurtainWallPanels, BuiltInCategory.OST_CurtainWallMullions]
	elementi = FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
	risultato = []
	for elem in elementi:
		if isinstance(elem, FamilyInstance) and elem.Category.BuiltInCategory not in categorie_escluse:
			if hasattr(elem.Host, "CurtainGrid") and elem.Host.CurtainGrid: 
				continue
			else:
				risultato.append(elem)
	return risultato

# Funzione: verifica posizione rispetto al livello host, con controllo esplicito su Z negativa
def verifica_elementi(elementi, livelli_arc, livelli_str):
	livelli_arc_id = [l.Id for l in livelli_arc]
	livelli_str_id = [l.Id for l in livelli_str]


	extra = []
	Risultato = []
	for elem in elementi:
		z = calcola_altezza(elem)
		host = elem.Host
		
		if isinstance(host, Level):
			host_id = host.Id
		elif hasattr(host, "LevelId"):
			host_id = host.LevelId
		elif not host and elem.Category.BuiltInCategory == BuiltInCategory.OST_GenericModel:
			host_id = elem.LevelId
		else:
			host_id = elem.LevelId
		lista_corrente = None
		livelli_id_corrente = []

		if host_id in livelli_arc_id:
			lista_corrente = livelli_arc
			livelli_id_corrente = livelli_arc_id
		elif host_id in livelli_str_id:
			lista_corrente = livelli_str
			livelli_id_corrente = livelli_str_id
		else:
			extra.append((elem, z))
			continue
		Risultato.append([elem,z,verifica_elevazione_superiore(z, lista_corrente, livelli_id_corrente.index(host_id))])

		"""
		if z < 0:
			verifiche.append(("Oggetto con offset negativo", elem, z))
		elif z > quota_superiore:
			verifiche.append(("Oggetto sopra l'interpiano", elem, z))
		else:
			verifiche.append(("Oggetto conforme", elem, z))
		"""
	return Risultato


Elementi = recupera_elementi(doc)
livelli_arc, livelli_str, livelli_errati = recupera_livelli(doc)


output.print_md("---")
stampa_livelli("AR", livelli_arc)
stampa_livelli("ST", livelli_str)

if livelli_errati:
	output.print_md("***Livelli con nome errato:***")
	for l in livelli_errati:
		output.print_md("**{}** : {}m".format(l.Name, converti_unita(l.Elevation)))
output.print_md("---")


elementi_posizionati = recupera_elementi(doc)

RisultatiAnalisi = verifica_elementi(elementi_posizionati, livelli_arc, livelli_str)

DataTable = []
for elem, z, (stato, simbolo,messaggio) in RisultatiAnalisi:
	DataTable.append([elem.Category.Name,output.linkify(elem.Id),z,simbolo,messaggio])
	if simbolo == ":cross_mark:":
		OFFSET_CSV_DATA.append([elem.Category.Name,elem.Id,messaggio,z,0])
	else:
		OFFSET_CSV_DATA.append([elem.Category.Name,elem.Id,messaggio,z,1])

"""

# DEFINIZIONE COLORI PER TROUBLESHOOTING
rosso = Color(255, 0, 0)
verde = Color(0, 255, 0)
blu = Color(0,0,255)
#viola = Color(150,0,150)
#arancio = Color(150,150,0)
#SolidPattern = FilteredElementCollector(doc).OfClass(FillPatternElement).FirstElement()


Override_ALTO = OverrideGraphicSettings()
Override_ALTO.SetProjectionLineColor(rosso)
#
Override_CORRETTO = OverrideGraphicSettings()
Override_CORRETTO.SetProjectionLineColor(verde)
#
Override_BASSO = OverrideGraphicSettings()
Override_BASSO.SetProjectionLineColor(blu)
#




# AVVIO VERIFICA
AllElements = FilteredElementCollector(doc,doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
Verifica = []
DataTable = []


t.Start()

for single_element in AllElements:
	if single_element.Category and single_element.Category.BuiltInCategory == BuiltInCategory.OST_Cameras:
		continue
	Famiglia = EstraiFamigliaOggetto(single_element)
	Livello_Oggetto = single_element.LevelId
	if Livello_Oggetto.Value > 0:
		Range = RegroupInfo[Livello_Oggetto][1]
	else:
		Livello_Oggetto = single_element.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)

	# USO LA BOUNDING BOX PER LA VERIFICA
	BoundingBoxElemento = single_element.get_BoundingBox(None)
	if BoundingBoxElemento:
		Min_Elev = ConvertiUnita(BoundingBoxElemento.Min.Z)
		Max_Elev = ConvertiUnita(BoundingBoxElemento.Max.Z)
		Mid_Elev = (Max_Elev+Min_Elev)/2
		
	# VERIFICO LA CATEGORIA DELL'ELEMENTO
		if single_element.Category.BuiltInCategory.ToString() in mep_built_in_categories:
			PuntoCalcolo = Min_Elev
		else:
			if single_element.Category.BuiltInCategory.ToString() =="OST_Floors" and single_element.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL).AsInteger() == 1:
				PuntoCalcolo = Max_Elev
			else:
				PuntoCalcolo = Mid_Elev
		if single_element.Category:
			if MathComparisonUtils.IsAlmostEqual(PuntoCalcolo, Range[0]):
						aview.SetElementOverrides(single_element.Id, Override_CORRETTO)
						Verifica.append(["Ok", single_element, single_element.Id])
			elif PuntoCalcolo > Range[1]:
				aview.SetElementOverrides(single_element.Id, Override_ALTO)
				Verifica.append(["Troppo Alto", single_element, single_element.Id])
				DataTable.append([Famiglia,output.linkify(single_element.Id),single_element.Category.Name,"Troppo Alto"])
				OFFSET_CSV_DATA.append([Famiglia,single_element.Id,single_element.Category.Name,"Troppo Alto",0])
			elif PuntoCalcolo < Range[0]:
				aview.SetElementOverrides(single_element.Id, Override_BASSO)
				Verifica.append(["Troppo Basso", single_element, single_element.Id])
				DataTable.append([Famiglia,output.linkify(single_element.Id),single_element.Category.Name,"Troppo Basso"])
				OFFSET_CSV_DATA.append([Famiglia,single_element.Id,single_element.Category.Name,"Troppo Basso",0])
			else:
				aview.SetElementOverrides(single_element.Id, Override_CORRETTO)
				Verifica.append(["Ok", single_element, single_element.Id])
				OFFSET_CSV_DATA.append([Famiglia,single_element.Id,single_element.Category.Name,"Verificato",1])
				#DataTable.append([single_element.get_Parameter(ELEM_FAMILY_AND_TYPE_PARAM).AsValueString(),single_element.Category.Name,"Posizionamento Corretto"])
	else:
		Verifica.append(["No BoundingBox", single_element, single_element.Id])
		#OFFSET_CSV_DATA.append([Famiglia,single_element.Id,single_element.Category,"ELEMENTO NON VERIFICABILE",0])

t.Commit()
"""



# CREAZIONE DELLA VISTA DI OUTPUT
if len(DataTable) != 0:
	OrderedData = sorted(DataTable, key=lambda status: status[-1])

	output = pyrevit.output.get_output()
	output.print_md("# Verifica Offset Da Livello Elementi")
	output.print_md("---")
	output.print_table(table_data = OrderedData, columns = ["Categoria Elemento", "ID Elemento","Elevazione", "Stato"], formats = ["","","",""])
	output.print_md("---")
else:
	output.print_md(":white_heavy_check_mark: **Tutti gli elementi sono inseriti correttamente.** :white_heavy_check_mark:")




###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
	return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si", "No"]
Scelta = pyrevit.forms.CommandSwitchWindow.show(ops, message="Esportare file CSV ?")
if Scelta == "Si":
	folder = pyrevit.forms.pick_folder()
	if folder:
		parameter_csv_path = os.path.join(folder, "13_ElementOffset_Data.csv")
		# Use codecs to open the file with UTF-8 encoding
		with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
			writer = csv.writer(file)
			writer.writerows(OFFSET_CSV_DATA)
		if VerificaTotale(OFFSET_CSV_DATA):
			pass

		else:
			pass
			


