# -*- coding: utf-8 -*-
""" Verifica la corretta collocazione altimetrica degli elementi rispetto al livello di host  """
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
OFFSET_CSV_DATA.append(["Famiglia e Tipo","ID Elemento","Categoria","Verifica","Stato"])

# INPUT DA PARTE DELL'UTENTE DEI VALORI DI VERIFICA
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox,Separator, Button, CheckBox)
components = [
			Label('Spessore Architettonico:'),
			TextBox('sp_arc', Text="0.30"),
			Label('Spessore Strutturale:'),
			TextBox('sp_str', Text="0.20"),
			Button('Select'),
			]
form = FlexForm('Definisci gli spessori di calcolo', components)
form.show()

if not form.values:
	script.exit()
	
ValoriUtente = form.values

# CATEGORIE MEP

mep_built_in_categories = [
	"OST_ElectricalFixtures",
	"OST_LightingFixtures",
	"OST_LightingDevices",
	"OST_CableTray",
	"OST_CableTrayFitting",
	"OST_Conduit",
	"OST_ConduitFitting",
	"OST_DataDevices",
	"OST_FireAlarmDevices",
	"OST_NurseCallDevices",
	"OST_SecurityDevices",
	"OST_TelephoneDevices",
	"OST_CommunicationDevices",
	"OST_DuctCurves",
	"OST_DuctFitting",
	"OST_DuctAccessory",
	"OST_DuctInsulations",
	"OST_DuctLinings",
	"OST_DuctSystem",
	"OST_MechanicalEquipment",
	"OST_PipeCurves",
	"OST_PipeFitting",
	"OST_PipeAccessory",
	"OST_PlumbingFixtures",
	"OST_PipeInsulations",
	"OST_Sprinklers",
	"OST_PipingSystem",
	"OST_DuctTerminal"
]
structural_built_in_categories = [
	"OST_StructuralColumns",
	"OST_StructuralFraming",
]
# CONVERTO STRINGHE IN ATTRIBUTI DI BUILTINCATEGORY
mep_built_in_category_enums = [getattr(BuiltInCategory, category) for category in mep_built_in_categories]

def ConvertiUnita(valore):
	Valore_Decimale = Decimal(UnitUtils.Convert(valore,UnitTypeId.Feet,UnitTypeId.Meters))
	getcontext().prec = 4
	#Valore_Raffinato = Valore_Decimale.quantize(Precisione_Decimale, rounding=ROUND_DOWN)
	return Valore_Decimale

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

# SPESSORI ARC E STR A SCELTA DELL'UTENTE

Spessore_ARC = float(ValoriUtente['sp_arc'])
Spessore_STR = float(ValoriUtente['sp_str'])
Spessore_Totale = Spessore_ARC + Spessore_STR

output.print_md("# Verifica Offset Elementi")
output.print_md("---")

# GESTIONE LIVELLI E SORTING PER CONSEQUENZIALITA (BYPASSA EVENTUALI ERRORI IN ORDINE DI CREAZIONE)

Livelli = FilteredElementCollector(doc).OfClass(Level).ToElements()
CoppiaLivelliAltimetria = [list((Livello.Id,ConvertiUnita(Livello.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble()))) for Livello in Livelli]
LivelliOrdinati = sorted(CoppiaLivelliAltimetria, key=lambda altezza: altezza[1])
##### CALCOLO I RANGE "LIMITE"
RangeCompetenza = []
RangeRelativo = []

for Index,Livello in enumerate(LivelliOrdinati):

	try:
		LimiteInferiore = ConvertiUnita(doc.GetElement(Livello[0]).get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble())-Decimal(Spessore_ARC)
		LimiteSuperiore = ConvertiUnita(doc.GetElement(LivelliOrdinati[Index+1][0]).get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble())-Decimal(Spessore_Totale)
		Differenza = LimiteInferiore-LimiteSuperiore
		RangeCompetenza.append([LimiteInferiore,LimiteSuperiore])
		RangeRelativo.append(round(Differenza,2))

	except:
		LimiteInferiore = ConvertiUnita(doc.GetElement(Livello[0]).get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble())-Decimal(Spessore_ARC)
		#LimiteSuperiore = ConvertiUnita(doc.GetElement(LivelliOrdinati[Index+1][0]).get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble())-Decimal(Spessore_Totale)
		RangeCompetenza.append([LimiteInferiore,1000])
		RangeRelativo.append([100-LimiteInferiore])


	
##### DEFINIZIONE DIZIONARIO INFO
RegroupInfo = {}
for Livello,Relativo,Range in zip(LivelliOrdinati,RangeRelativo,RangeCompetenza):
	RegroupInfo[Livello[0]]=[Relativo,Range]


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

#Override_TROPPOLUNGO = OverrideGraphicSettings()
#Override_TROPPOLUNGO.SetSurfaceForegroundPatternId(SolidPattern.Id)
#Override_TROPPOLUNGO.SetSurfaceForegroundPatternColor(viola)


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

# CREAZIONE DELLA VISTA DI OUTPUT
if len(DataTable) != 0:
	OrderedData = sorted(DataTable, key=lambda status: status[-1])

	output = pyrevit.output.get_output()
	output.print_md("# Verifica Offset Da Livello Elementi")
	output.print_md("---")
	output.print_md("***Spessore Architettonico: {}m Spessore Strutturale : {}m***".format(Spessore_ARC,Spessore_STR))
	output.print_md("---")
	output.print_table(table_data = OrderedData, columns = ["Nome Elemento", "ID Elemento","Categoria", "Stato"], formats = ["","","",""])
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
		if VerificaTotale(OFFSET_CSV_DATA):
			OFFSET_CSV_DATA = []
			OFFSET_CSV_DATA.append(["Nome Verifica","Stato"])
			OFFSET_CSV_DATA.append(["Regole di modellazione - Elementi posizionati correttamente.",1])
			parameter_csv_path = os.path.join(folder, "13_XX_ElementOffset_Data.csv")
			# Use codecs to open the file with UTF-8 encoding
			with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerows(OFFSET_CSV_DATA)
		else:
			parameter_csv_path = os.path.join(folder, "13_XX_ElementOffset_Data.csv")
			# Use codecs to open the file with UTF-8 encoding
			with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerows(OFFSET_CSV_DATA)



