# -*- coding: utf-8 -*-

""" Verifica le unita di misura utilizzate nel progetto aperto  """

__author__ = 'Roberto Dolfini'
__title__ = 'Verifica Unita Misura'

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
import codecs
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from System.Collections.Generic import *


##############################################################
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument			   
app   = __revit__.Application	  
aview = doc.ActiveView
output = pyrevit.output.get_output()


##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
#script_dir = os.path.dirname(__file__)
#parent_dir = os.path.abspath(os.path.join(script_dir, '..','..','000_Raccolta CSV di controllo','12_CSV_Nomenclatura Livelli.csv'))


# CREAZIONE LISTE DI OUTPUT DATA
VERIFICAUNITA_CSV_OUTPUT = []
VERIFICAUNITA_CSV_OUTPUT.append(["Unita di misura","Formato corrente","Stato"])

output = pyrevit.output.get_output()

Project_Units = doc.GetUnits()

ValoriDaControllare = [
	SpecTypeId.Length,
	SpecTypeId.Area,
	SpecTypeId.Volume,
	SpecTypeId.Distance
]

DizionarioCorrezione = {
	"meters":"Metri",
	"squareMeters":"Metri quadrati",
	"cubicMeters":"Metri cubi",
	"centimeters":"Centimetri",
	"millimeters":"Millimetri"
}

DizionarioUnita = {
	"LENGTH":"Lunghezza",
	"AREA":"Area",
	"VOLUME":"Volume",
	"DISTANCE":"Distanza"
}

Format_Options = []
for Valore in ValoriDaControllare:
	Format_Options.append(Project_Units.GetFormatOptions(Valore).GetUnitTypeId().TypeId.ToString())

DataTable = []
for Val,Tipo in zip(Format_Options,ValoriDaControllare):

	try:
		Valore_Unita = DizionarioCorrezione[Val.split(":")[1].split("-")[0]]
	except:
		Valore_Unita = Val.split(":")[1].split("-")[0]
	try:
		Valore_Tipo = DizionarioUnita[Tipo.TypeId.ToString().split(":")[1].split("-")[0].upper()]
	except:
		Valore_Tipo = Tipo.TypeId.ToString().split(":")[1].split("-")[0].upper()
			
	if "metri" in Valore_Unita.lower():
		DataTable.append([Valore_Tipo,Valore_Unita,":white_heavy_check_mark:"])
		VERIFICAUNITA_CSV_OUTPUT.append([Valore_Tipo,Valore_Unita,1])
	else:
		DataTable.append([Valore_Tipo,Valore_Unita,":cross_mark:"])
		VERIFICAUNITA_CSV_OUTPUT.append([Valore_Tipo,Valore_Unita,0])

output.print_md("# Verifica unita Di Progetto")
output.print_md("---")
output.print_table(table_data = DataTable,columns = ["Unita di misura","Formato corrente","Sistema Metrico"],formats = ["","",""])


###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
	return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")

if Scelta == "Si":
	folder = pyrevit.forms.pick_folder()
	if folder:
		copymonitor_csv_path = os.path.join(folder, "15_UnitaProgetto_Data.csv")
	with codecs.open(copymonitor_csv_path, mode='w', encoding='utf-8') as file:
		writer = csv.writer(file)
		writer.writerows(VERIFICAUNITA_CSV_OUTPUT)
	if VerificaTotale(VERIFICAUNITA_CSV_OUTPUT):
		"""
		VERIFICAUNITA_CSV_OUTPUT = []
		VERIFICAUNITA_CSV_OUTPUT.append(["Nome Verifica","Stato"])
		VERIFICAUNITA_CSV_OUTPUT.append(["Georeferenziazione e Orientamento - CopyMonitor correttamente effettuato.",1])
		copymonitor_csv_path = os.path.join(folder, "07_02_CopyMonitorReport_Data.csv")
		with codecs.open(copymonitor_csv_path, mode='w', encoding='utf-8') as file:
		writer = csv.writer(file)
		writer.writerows(VERIFICAUNITA_CSV_OUTPUT)
			"""
	else:
		pass

