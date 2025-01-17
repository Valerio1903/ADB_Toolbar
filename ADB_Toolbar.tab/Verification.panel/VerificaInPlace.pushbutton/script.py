# -*- coding: utf-8 -*-

""" Verifica la presenza di famiglie model in place nel progetto """

__author__ = 'Roberto Dolfini'
__title__ = 'Verifica Famiglie InPlace'

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
import codecs
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *
#PER IL FLOATING POINT
from decimal import Decimal, ROUND_DOWN, getcontext

##############################################################
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument               
app   = __revit__.Application      
aview = doc.ActiveView
output = pyrevit.output.get_output()

t = Transaction(doc, "Verifica Fase")

##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
#script_dir = os.path.dirname(__file__)
#parent_dir = os.path.abspath(os.path.join(script_dir, '..','..','000_Raccolta CSV di controllo','12_CSV_Nomenclatura Livelli.csv'))


# CREAZIONE LISTE DI OUTPUT DATA
FAMIGLIEINPLACE_CSV_OUTPUT = []
FAMIGLIEINPLACE_CSV_OUTPUT.append(["Categoria","Famiglia","Id Elemento","Stato"])

output = pyrevit.output.get_output()

CollectorProgetto = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()



DataTable = []

for element in CollectorProgetto:
    if isinstance(element, FamilyInstance) and element.Symbol.Family.IsInPlace:
        DataTable.append([element.Symbol.Family.FamilyCategory.Name,element.Symbol.Family.Name,element.Id])
        FAMIGLIEINPLACE_CSV_OUTPUT.append([element.Symbol.Family.FamilyCategory.Name,element.Symbol.Family.Name,element.Id,0])
		
    elif isinstance(element, FamilyInstance) and not element.Symbol.Family.IsInPlace:
        FAMIGLIEINPLACE_CSV_OUTPUT.append([element.Symbol.Family.FamilyCategory.Name,element.Symbol.Family.Name,element.Id,1])
		
output.print_md("# Verifica Famiglie In-Place")
output.print_md("---")
if DataTable:
	output.print_table(table_data = DataTable,columns = ["Categoria","Famiglia","Id Elemento"],formats = ["","",""])
else:
	output.print_md(":white_heavy_check_mark: **Nessuna famiglia in-place presente nel progetto** :white_heavy_check_mark: ")



###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
	return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")

if Scelta == "Si":
	folder = pyrevit.forms.pick_folder()
	if folder:
		
		if VerificaTotale(FAMIGLIEINPLACE_CSV_OUTPUT):
			pass
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
			copymonitor_csv_path = os.path.join(folder, "13_FamiglieInPlace_Data.csv")
			with codecs.open(copymonitor_csv_path, mode='w', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerows(FAMIGLIEINPLACE_CSV_OUTPUT)
