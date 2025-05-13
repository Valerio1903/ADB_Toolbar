# -*- coding: utf-8 -*-
""" Verifica la presenza di workset all'interno del modello di consegna.  """

__author__ = 'Roberto Dolfini'
__title__ = 'Verifica Presenza\nWorkset'
import codecs
import re
import unicodedata
import pyrevit
from pyrevit import *
from pyrevit import forms, script
import sys
import clr
import System
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import csv
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType
import math
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

##############################################################
doc   = __revit__.ActiveUIDocument.Document  #type: Document
uidoc = __revit__.ActiveUIDocument						   
app   = __revit__.Application		
output = pyrevit.output.get_output()
aview = doc.ActiveView
##############################################################

# CREAZIONE LISTE DI OUTPUT DATA
WORKSET_CSV_OUTPUT = []
WORKSET_CSV_OUTPUT.append(["Nome Workset","Workset Id", "Unique ID","Kind","n° elementi","Stato"])

# Trova il percorso della cartella corrente (dove si trova anche lettore.py)
current_folder = os.path.dirname(__file__)
lettore_script = os.path.join(current_folder, "lettore.py")
"""
# Percorso all'interprete Python 3 (adatta in base alla tua installazione)
python_exe = r"C:\Users\2Dto6D\AppData\Local\Programs\Python\Python311\python.exe"
"""

output.print_md("# Verifica Workset nel modello")
output.print_md("---")

WorksetCollector = FilteredWorksetCollector(doc).ToWorksets()

NomeWorkset, IdWorkset, UniqueIdWorkset, KindWorkset, nElementi = [], [], [], [], []
for wks in WorksetCollector:
    NomeWorkset.append(wks.Name)
    IdWorkset.append(wks.Id)
    UniqueIdWorkset.append(wks.UniqueId)
    KindWorkset.append(wks.Kind)

DataTable = []
for n,id,ui,kind in zip(NomeWorkset,IdWorkset,UniqueIdWorkset,KindWorkset):
    if kind != WorksetKind.UserWorkset:
        continue
    elementWorksetFilter = ElementWorksetFilter(id)
    CollectorElementi = FilteredElementCollector(doc).WherePasses(elementWorksetFilter).ToElements()
    nElementi.append(len(CollectorElementi))
    DataTable.append([n, id, ui, kind, len(CollectorElementi)])

# PRINT TABELLA DI VERIFICA
output.print_table(table_data = DataTable, columns = ["Nome Workset","WorksetId", "Unique ID","Kind","n° elementi"])

if len(DataTable) == 1:
    output.print_md("###:white_heavy_check_mark: Rilevato un solo workset nel modello.")
else:
    output.print_md("###:cross_mark: Rilevati più workset nel modello. :cross_mark:")

# OUTPUT CSV
if len(DataTable) == 1:
    WORKSET_CSV_OUTPUT.append([DataTable[0][0], DataTable[0][1], DataTable[0][2], DataTable[0][3], DataTable[0][4],1])
else:
    for Data in DataTable:
        WORKSET_CSV_OUTPUT.append([Data[0], Data[1], Data[2], Data[3], Data[4], 0])


###OPZIONI ESPORTAZIONE

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")

if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        copymonitor_csv_path = os.path.join(folder, "10_PresenzaWorkset.csv")
        with codecs.open(copymonitor_csv_path, mode='w', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(WORKSET_CSV_OUTPUT)
