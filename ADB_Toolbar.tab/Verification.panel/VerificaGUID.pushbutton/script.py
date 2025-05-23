# -*- coding: utf-8 -*-
""" Verifica la presenza di parametri condivisi e, se trovati, verifica la corrispondenza dei GUID  """
__author__ = 'Roberto Dolfini'
__title__ = 'Verifica Parametri\nCondivisi'

import codecs
import re
import unicodedata
import pyrevit
from pyrevit import *
import clr
import System
from pyrevit import forms, script
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import csv
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType
import codecs
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *


##############################################################
doc   = __revit__.ActiveUIDocument.Document  #type: Document
uidoc = __revit__.ActiveUIDocument						   
app   = __revit__.Application		
##############################################################  

# CREAZIONE LISTE DI OUTPUT DATA
PARAMETERS_CSV_DATA = []

# CREAZIONE DELLA VISTA DI OUTPUT
output = pyrevit.output.get_output()
output.print_md("# Verifica Parametri Condivisi")
output.print_md("---")

# ESTRAGGO LE INFO DAL FILE ESTERNO 
shared_params_file = app.OpenSharedParameterFile()

if not shared_params_file:
	output.print_md(":cross_mark: ***FILE DEI PARAMETRI CONDIVISI NON TROVATO ALL'INTERNO DEL PROGETTO***")
	PARAMETERS_CSV_DATA.append(["FILE DEI PARAMETRI CONDIVISI NON TROVATO ALL'INTERNO DEL PROGETTO",0])
else:
	ParametersFromFile = []
	ParametersFromFileClean = []
	for group in shared_params_file.Groups:
		Regroup = []
		for definition in group.Definitions:
			Regroup.append([group.Name,definition.Name,str(definition.GUID)])
			ParametersFromFile.append([group.Name,definition.Name,str(definition.GUID)])
		ParametersFromFileClean.append(Regroup)
		
# ESTRAGGO I PARAMETRI PRESENTI NEL FILE
ProjectBindings = doc.ParameterBindings
FWDIterator = ProjectBindings.ForwardIterator()
ParametersFromRVT = []
while FWDIterator.MoveNext():
	definition = FWDIterator.Key
	if isinstance(FWDIterator.Current, TypeBinding):
		binding = "Parametro di Tipo"
	elif isinstance(FWDIterator.Current, InstanceBinding):
		binding = "Parametro di Istanza"
	else:
		binding = "Parametro Sconosciuto"
	param_name = definition.Name
	param_id = definition.Id
	try:
		is_shared = doc.GetElement(param_id).GuidValue
	except:
		is_shared = "Parametro non condiviso"
	ParametersFromRVT.append([param_name,str(param_id),str(is_shared)])
if ParametersFromRVT:
	
	output.print_md("## Parametri Condivisi da File")
	output.print_table(table_data = ParametersFromFile,columns = ["Nome Gruppo","Nome Parametro", "GUID"], formats = ["","",""])
	output.print_md("## Parametri Condivisi da Progetto")
	output.print_table(table_data = ParametersFromRVT,columns = ["Nome Parametro","Id Parametro","Condiviso ?"], formats = ["","",""])
	output.print_md("---")

	OutputResult = []
	PARAMETERS_CSV_DATA.append(["Nome Parametro","GUID Parametro","Stato"])
	for ProjectParameter in ParametersFromRVT:
		FOUND = False
		for SharedList in ParametersFromFileClean:
			for SharedParameter in SharedList:
				if ProjectParameter[-1] == SharedParameter[-1]:
					FOUND = True

		if FOUND:
			Name = ProjectParameter[0] 
			GUID = ProjectParameter[-1]
			Status = "1"
			Symbol = ":white_heavy_check_mark:"
			OutputResult.append([Name,GUID,Symbol])
			PARAMETERS_CSV_DATA.append([Name,GUID,Status])


		else:
			Name = str(ProjectParameter[0])
			GUID = str(ProjectParameter[-1])
			ColorName = "<b style=color:red;'>{}</b><br>".format(str(ProjectParameter[0]))
			ColorGUID = "<b style=color:red;'>{}</b><br>".format(str(ProjectParameter[-1]))
			Status = "0"
			Symbol = ":cross_mark:"
			OutputResult.append([ColorName,ColorGUID,Symbol])
			PARAMETERS_CSV_DATA.append([Name,GUID,Status])
			
	output.print_md("## Verifica Corrispondenza")
	output.print_table(table_data = OutputResult,columns = ["Nome Parametro","GUID","Condiviso ?"], formats = ["","",""])			
else:
	output.print_md(":cross_mark: **ATTENZIONE, NON CI SONO PARAMETRI DI PROGETTO NEL FILE**")
	PARAMETERS_CSV_DATA.append(["ATTENZIONE, NON CI SONO PARAMETRI DI PROGETTO NEL FILE",0])

###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
	return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))


ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
	folder = pyrevit.forms.pick_folder()
	if folder:
		parameter_csv_path = os.path.join(folder, "10_SharedParameters_Data.csv")
		with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
			writer = csv.writer(file)
			writer.writerows(PARAMETERS_CSV_DATA)
		if VerificaTotale(PARAMETERS_CSV_DATA):
			pass
			""" PER ORA RIMOSSO IN ATTESA DI SPECIFICHE
			PARAMETERS_CSV_DATA = []
			PARAMETERS_CSV_DATA.append(["Nome Verifica","Stato"])
			PARAMETERS_CSV_DATA.append(["Integrità e pulizia file - Parametri correttamente valorizzati.",1])
			parameter_csv_path = os.path.join(folder, "10_01_SharedParameters_Data.csv")
			with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerows(PARAMETERS_CSV_DATA)
			"""
		else:
			pass