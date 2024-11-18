""" Verifica che gli elementi strutturali non eccedano la lunghezza massima di 5 metri. """

__title__ = 'Verifica Lunghezza\nElementi Strutturali'

import codecs
import re
import unicodedata
import pyrevit
from pyrevit import forms, script
import math
import clr
import os
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

t = Transaction(doc, "Verifica Strutturale")

##############################################################

# VERIFICA PRELIMINARE DI SELEZIONE ELEMENTI
if not uidoc.Selection.GetElementIds():
	forms.alert("Nessun elemento selezionato tenere gli elementi da analizzare in selezione attiva.")
	script.exit()

# CREAZIONE LISTE DI OUTPUT DATA
STRUCTURAL_ERROR_CSV_OUTPUT = []
STRUCTURAL_ERROR_CSV_OUTPUT.append(["Famiglia e Tipo","ID Elemento","Verifica","Stato"])

structural_categories = [
"OST_StructuralTrussStickSymbols",
"OST_StructuralTrussHiddenLines",
"OST_StructuralFramingSystemHiddenLines_Obsolete",
"OST_StructuralTendonTags",
"OST_StructuralTendonHiddenLines",
"OST_StructuralTendons",
"OST_StructuralBracePlanReps",
"OST_StructuralAnnotations",
"OST_StructuralConnectionHandlerTags_Deprecated",
"OST_StructuralFoundationTags",
"OST_StructuralColumnTags",
"OST_StructuralFramingTags",
"OST_StructuralStiffenerHiddenLines",
"OST_StructuralColumnLocationLine",
"OST_StructuralFramingLocationLine",
"OST_StructuralStiffenerTags",
"OST_StructuralStiffener",
"OST_StructuralTruss",
"OST_StructuralColumnStickSymbols",
"OST_HiddenStructuralColumnLines",
"OST_StructuralColumns",
"OST_HiddenStructuralFramingLines",
"OST_StructuralFramingSystem",
"OST_StructuralFramingOther",
"OST_StructuralFraming",
"OST_HiddenStructuralFoundationLines",
"OST_StructuralFoundation",
"OST_StructuralFramingOpening",
"OST_HiddenStructuralConnectionLines_Deprecated",
"OST_StructuralConnectionHandler_Deprecated"
]

# CONVERTO STRINGHE IN ATTRIBUTI DI BUILTINCATEGORY
str_built_in_category_enums = [getattr(BuiltInCategory, category) for category in structural_categories]

def ConvertiUnita(valore):
	Valore_Decimale = Decimal(UnitUtils.Convert(valore,UnitTypeId.Feet,UnitTypeId.Meters))
	getcontext().prec = 4
	#Valore_Raffinato = Valore_Decimale.quantize(Precisione_Decimale, rounding=ROUND_DOWN)
	return Valore_Decimale

def EstraiFamigliaOggetto(oggetto):
	try:
		return oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
	except:
		try:
			return oggetto.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsValueString()
		except:
			return "NON TROVATO"
######################################################################################
AllElementsIds = uidoc.Selection.GetElementIds()

# DEFINIZIONE COLORI PER TROUBLESHOOTING
rosso = Color(255, 0, 0)
verde = Color(0, 255, 0)
SolidPattern = FilteredElementCollector(doc).OfClass(FillPatternElement).FirstElement()

#
Override_CORRETTO = OverrideGraphicSettings()
Override_CORRETTO.SetProjectionLineColor(verde)
Override_CORRETTO.SetSurfaceForegroundPatternId(SolidPattern.Id)
#
Override_ERRATO = OverrideGraphicSettings()
Override_ERRATO.SetProjectionLineColor(rosso)
Override_ERRATO.SetSurfaceForegroundPatternId(SolidPattern.Id)
#

# AVVIO VERIFICA
AllElements = [doc.GetElement(x) for x in AllElementsIds]
Verifica = []
DataTable = []

t.Start()
for single_element in AllElements:
	if single_element.Category.BuiltInCategory.ToString() in structural_categories:
			Famiglia = EstraiFamigliaOggetto(single_element)
			Lunghezza = ConvertiUnita(single_element.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsDouble())
   
			if 	Lunghezza > 5:		
				DataTable.append([Famiglia,output.linkify(single_element.Id),single_element.Category.Name,"Troppo Lungo"])
				Verifica.append(["Troppo Lungo", single_element, single_element.Id])
				STRUCTURAL_ERROR_CSV_OUTPUT.append([Famiglia,single_element.Id,single_element.Category.Name,"Troppo Lungo",0])
				aview.SetElementOverrides(single_element.Id, Override_ERRATO)
			else:
				Verifica.append(["Ok", single_element, single_element.Id])
				STRUCTURAL_ERROR_CSV_OUTPUT.append([Famiglia,single_element.Id,single_element.Category.Name,"Verificato",1])
				aview.SetElementOverrides(single_element.Id, Override_CORRETTO)
t.Commit()

# CREAZIONE DELLA VISTA DI OUTPUT
OrderedData = sorted(DataTable, key=lambda status: status[-1])

output = pyrevit.output.get_output()
output.print_md("# Verifica Lunghezza Elementi Strutturali")
output.print_md("---")
output.print_table(table_data = OrderedData, columns = ["Nome Elemento", "ID Elemento","Categoria", "Stato"], formats = ["","","",""])
output.print_md("---")


###OPZIONI ESPORTAZIONE
ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
	folder = pyrevit.forms.pick_folder()
	if folder:
		parameter_csv_path = os.path.join(folder, "StructuralError_Data.csv")
		with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
			writer = csv.writer(file)
			writer.writerows(STRUCTURAL_ERROR_CSV_OUTPUT)
