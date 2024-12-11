# -*- coding: utf-8 -*-

""" Verifica la corretta nomenclatura delle fasi presenti  """
__title__ = 'Check Nomenclatura Fasi'

######################################

import Autodesk.Revit
import Autodesk.Revit.DB
import Autodesk.Revit.DB.Architecture
import pyrevit
from pyrevit import forms, script
import codecs
import re
from re import *
from Autodesk.Revit.DB import BuiltInCategory, Category, CategoryType
import System
from System import Enum
from System.Collections.Generic import Dictionary
import unicodedata
import math
import clr
import os
import sys
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import csv
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType
import time
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

#t = Transaction(doc, "Verifica Nomenclatura Griglie e Livelli")
##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
script_dir = os.path.dirname(__file__)
parent_dir = os.path.abspath(os.path.join(script_dir, '..','Raccolta CSV di controllo','Database_StrutturaNomenclaturaLivelli.csv'))

#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output()

# CREAZIONE LISTE DI OUTPUT DATA
PHASE_NAMING_CSV_OUTPUT = []
PHASE_NAMING_CSV_OUTPUT.append(["Nome Elemento","ID Elemento","Stato"])
##############################################################

# COLLECTOR GRIGLIE E LIVELLI
Collector_Fasi = FilteredElementCollector(doc).OfClass(Phase).ToElements()

output.print_md("# Verifica Nomenclatura Fasi")
output.print_md("---")


for Fase in Collector_Fasi:
    Nome = Fase.Name
    if len(Fase.Name.split("_")) != 5:
        VERIFICA = "Nome incorretto"
        SIMBOLO = ":cross_mark:"
        VALUE = 0
    else:
        SuddivisioneNome = Nome.split("_")
        VERIFICA = "Codice disciplina errato"
        SIMBOLO = ":cross_mark:"
        VALUE = 0


"""
# ACCEDO AL CSV DI CONTROLLO E CREO IL DIZIONARIO DI VERIFICA
with codecs.open(parent_dir, 'r', 'utf-8-sig') as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        Chiave=row[0]
        Codifica=row[1].split("~")
        DizionarioDiVerifica[Chiave] = Codifica

VERIFICA = "NUMERO CAMPI ERRATO"
SIMBOLO = ":cross_mark:"


# Append results to DataTable for both errors and successes
DataTable.append([Livello.Category.Name, Nome, output.linkify((Livello.Id)), VERIFICA, SIMBOLO])
GRID_LEVELS_NAMING_CSV_OUTPUT.append([Livello.Category.Name, Livello.Name, Livello.Id, VALUE])

output.print_table(table_data = DataTable,title = "Verifica Nomenclatura Livelli", columns = ["Categoria","ID Elemento","Verifica"],formats = ["","",""])

###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        if VerificaTotale(GRID_LEVELS_NAMING_CSV_OUTPUT):
            GRID_LEVELS_NAMING_CSV_OUTPUT.append("Nome Verifica","Stato")
            GRID_LEVELS_NAMING_CSV_OUTPUT.append("Naming Convention - Nomenclatura Griglie e Livelli.",1)
        else:
            gridsandlevels_csv_path = os.path.join(folder, "12_XX_CoordinationReport_Data.csv")
            with codecs.open(gridsandlevels_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(GRID_LEVELS_NAMING_CSV_OUTPUT)

"""