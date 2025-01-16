# -*- coding: utf-8 -*-

""" Verifica l'assegnazione della fase corretta ai diversi elementi presenti in vista  """

__author__ = 'Roberto Dolfini'
__title__ = 'Verifica Nomenclatura Fasi'

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
doc   = __revit__.ActiveUIDocument.Document  #type: Document
uidoc = __revit__.ActiveUIDocument               
app   = __revit__.Application      
aview = doc.ActiveView
output = pyrevit.output.get_output()

t = Transaction(doc, "Verifica Fase")

##############################################################

#COLLOCAZIONE CSV DI CONTROLLO
script_dir = os.path.dirname(__file__)
parent_dir = os.path.abspath(os.path.join(script_dir, '..','..','000_Raccolta CSV di controllo','12_CSV_Nomenclatura Livelli.csv'))


# CREAZIONE LISTE DI OUTPUT DATA
VERIFICAFASE_CSV_OUTPUT = []
VERIFICAFASE_CSV_OUTPUT.append(["Nome Fase", "Verifica", "Stato"])

output = pyrevit.output.get_output()
output.print_md("# Verifica Nomenclatura Fasi")
output.print_md("---")

#VERIFICO LA PRESENZA DELLA FASE ALL'INTERNO DEL FILE

FasiDelDocumento = FilteredElementCollector(doc).OfClass(Phase)
FasiPresenti = [Fase.Name for Fase in FasiDelDocumento]

# ACCEDO AL CSV DI CONTROLLO E CREO IL DIZIONARIO DI VERIFICA
with codecs.open(parent_dir, 'r', 'utf-8-sig') as f:
    reader = csv.reader(f, delimiter=',')
    rows = list(reader)  # Convert to a list of rows
    if len(rows) >= 1:  # Ensure CSV has at least 2 rows
        FasiDatabase = rows[1]
    else:
        output.print_md(":warning: CSV file does not have enough rows for verification!")

DataTable = []

for FasePresente in FasiPresenti:
    VERIFICA = "Fase non presente nel database."
    SIMBOLO = ":cross_mark:"
    VALUE = 0 
    if FasePresente in FasiDatabase:
        VERIFICA = "Fase presente nel database."
        SIMBOLO = ":white_check_mark:"
        VALUE = 1
    VERIFICAFASE_CSV_OUTPUT.append([FasePresente,VERIFICA,VALUE])
    DataTable.append([FasePresente,VERIFICA,SIMBOLO])

# CREAZIONE DELLA VISTA DI OUTPUT
if len(DataTable) != 0:
    output.freeze()
    output.print_table(table_data = DataTable, columns = ["Fase","Verifica","Stato"], formats = ["","",""])
    output.unfreeze()

###OPZIONI ESPORTAZIONE
ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")
if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        parameter_csv_path = os.path.join(folder, "12_FaseErrata_Data.csv")
    with codecs.open(parameter_csv_path, mode='w', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(VERIFICAFASE_CSV_OUTPUT)

