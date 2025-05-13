# -*- coding: utf-8 -*-

""" Estrae la starting view dal progetto """

__author__ = 'Roberto Dolfini'
__title__ = 'Starting View Info'

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

##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
#script_dir = os.path.dirname(__file__)
#parent_dir = os.path.abspath(os.path.join(script_dir, '..','..','000_Raccolta CSV di controllo','12_CSV_Nomenclatura Livelli.csv'))


# CREAZIONE LISTE DI OUTPUT DATA
STARTINGVIEW_CSV_OUTPUT = []
STARTINGVIEW_CSV_OUTPUT.append(["Nome Starting View","Id Elemento","Presenza Starting View","Presenza Informazioni","Stato"])

output = pyrevit.output.get_output()

output.print_md("# Verifica Starting View")
output.print_md("---")

StartingViewSetup = StartingViewSettings.GetStartingViewSettings(doc)

StartingViewId = StartingViewSetup.ViewId
INFO = "Mancano informazioni, si rimanda all'allegato DD2_REPORT"

if StartingViewId.IntegerValue == -1:
	VALUE = 0
	SIMBOLO = ":cross_mark:"
	VERIFICA = "Starting View non impostata"
	StartingViewName = "Nessuna"
	output.print_md(":cross_mark: **Starting View non impostata** :cross_mark:")
	output.print_md("**{}**".format(INFO))
	STARTINGVIEW_CSV_OUTPUT.append([StartingViewName, StartingViewId,VERIFICA , INFO,0])

#! CORREGGERE L'OUTPUT DI INFO

else:
	VALUE = 1
	SIMBOLO = ":white_check_mark:"
	VERIFICA = "Starting View impostata"
	StartingViewName = doc.GetElement(StartingViewId).Name
	output.print_md("##:white_heavy_check_mark: Starting View impostata :white_heavy_check_mark:")
	output.print_md("**Starting View: {}**".format(StartingViewName))
	output.print_md("**Id Elemento: {}**".format(output.linkify(StartingViewId)))
	STARTINGVIEW_CSV_OUTPUT.append([StartingViewName, StartingViewId,VERIFICA,INFO,0])


ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")

SuddivisioneNomeFile = doc.Title.split("_")
DisciplinaFile =SuddivisioneNomeFile[4]+SuddivisioneNomeFile[6]
RevisioneFile = SuddivisioneNomeFile[7]
CodiceVerifica = "09"
NomeVerifica = "StartingView_Data.csv"

NomeFileExport = DisciplinaFile + "_" + CodiceVerifica + "_" + RevisioneFile + "_" + NomeVerifica

if Scelta == "Si":
	folder = pyrevit.forms.pick_folder()
	if folder:
		copymonitor_csv_path = os.path.join(folder, NomeFileExport)
		with codecs.open(copymonitor_csv_path, mode='w', encoding='utf-8') as file:
			writer = csv.writer(file)
			for row in STARTINGVIEW_CSV_OUTPUT:
				writer.writerow(row)
