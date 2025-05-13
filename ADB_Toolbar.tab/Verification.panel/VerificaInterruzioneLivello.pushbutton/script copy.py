# -*- coding: utf-8 -*-

""" Verifica che gli elementi siano correttamente vincolati tra livello e livello.  """

__author__ = 'Roberto Dolfini'
__title__ = 'Verifica Interruzione elementi tra livelli.'

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

#t = Transaction(doc, "Verifica Versione Files")

##############################################################

# CREAZIONE LISTE DI OUTPUT DATA
VINCOLOLIVELLI_CSV_OUTPUT = []
VINCOLOLIVELLI_CSV_OUTPUT.append(["Categoria","Id Elemento","Livello Base","Livello Superiore","Verifica","Stato"])


#ESTRAZIONE LIVELLI E RELATIVE INFO
LivelliDocumento = FilteredElementCollector(doc).OfClass(Level).ToElements()

DatiLivelli = []
for Livello in LivelliDocumento:
    DatiLivelli.append([Livello.Id,Livello.Elevation])

DatiLivelli.sort(key=lambda x: x[1])
#ESTRAZIONE SOLO ID DEI LIVELLI
IDLivelli = [x[0] for x in DatiLivelli]

#ESTRAZIONE MURI - COLONNE ED INFO
CatFilter = List[BuiltInCategory]()
CatFilter.Add(BuiltInCategory.OST_Walls)
CatFilter.Add(BuiltInCategory.OST_Columns)
CatFilter.Add(BuiltInCategory.OST_StructuralColumns)

FiltroCategorie = ElementMulticategoryFilter(CatFilter)

ElementiVerifica = FilteredElementCollector(doc).WherePasses(FiltroCategorie).WhereElementIsNotElementType().ToElements()

DataTable = []

for Elemento in ElementiVerifica:
    try:
        Base = Elemento.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
        Superiore = Elemento.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).AsElementId()
        NomeBase = Elemento.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsValueString()
        NomeSuperiore = Elemento.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).AsValueString()
    except:
        Base = Elemento.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsElementId()
        Superiore = Elemento.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId()
        NomeBase = Elemento.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsValueString()
        NomeSuperiore = Elemento.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsValueString()    

    if Superiore != ElementId(-1):
        if IDLivelli[IDLivelli.index(Base)+1] != Superiore:
            VALUE = 0
            VERIFICA = "Elemento non correttamente vincolato tra livelli."
            SIMBOLO = ":cross_mark:"
            if ":" in NomeSuperiore:
                StringaPulita = NomeSuperiore.split(":")[1]
            else:
                StringaPulita = NomeSuperiore
            DataTable.append([Elemento.Category.Name,Elemento.Id,NomeBase,StringaPulita,VERIFICA,SIMBOLO])
            VINCOLOLIVELLI_CSV_OUTPUT.append([Elemento.Category.Name,Elemento.Id,NomeBase,StringaPulita,VERIFICA,VALUE])
        else:
            VALUE = 1
            VERIFICA = "Correttamente vincolato tra livelli."
            SIMBOLO = ":white_heavy_check_mark:"
            if ":" in NomeSuperiore:
                StringaPulita = NomeSuperiore.split(":")[1]
            else:
                StringaPulita = NomeSuperiore
            DataTable.append([Elemento.Category.Name,Elemento.Id,NomeBase,StringaPulita,VERIFICA,SIMBOLO])
            VINCOLOLIVELLI_CSV_OUTPUT.append([Elemento.Category.Name,Elemento.Id,NomeBase,StringaPulita,VERIFICA,VALUE])
    else:
        VALUE = 0
        VERIFICA = "Non presenta vincolo superiore."
        SIMBOLO = ":cross_mark:"
        DataTable.append([Elemento.Category.Name,Elemento.Id,NomeBase,NomeSuperiore,VERIFICA,SIMBOLO])
        VINCOLOLIVELLI_CSV_OUTPUT.append([Elemento.Category.Name,Elemento.Id,NomeBase,NomeSuperiore,VERIFICA,VALUE])


output.print_md("# Verifica vincolo elementi verticali")
output.print_md("---")
output.freeze()
output.print_table(table_data = DataTable,columns = ["Categoria","Id Elemento","Livello Base","Livello Superiore","Verifica","Stato"],formats = ["","","","","",""])
output.unfreeze()

###OPZIONI ESPORTAZIONE
def VerificaTotale(lista):
    return all(sublist[-1] == 1 for sublist in lista if isinstance(sublist[-1], int))

ops = ["Si","No"]
Scelta = forms.CommandSwitchWindow.show(ops, message ="Esportare file CSV ?")

if Scelta == "Si":
    folder = pyrevit.forms.pick_folder()
    if folder:
        copymonitor_csv_path = os.path.join(folder, "13_VincoloTraLivelli_Data.csv")
        with codecs.open(copymonitor_csv_path, mode='w', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(VINCOLOLIVELLI_CSV_OUTPUT)
        if VerificaTotale(VINCOLOLIVELLI_CSV_OUTPUT):
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
            pass
