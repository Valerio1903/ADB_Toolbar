# -*- coding: utf-8 -*-

""" Verifica la corretta nomenclatura di griglie e livelli  """
__author__= 'Roberto Dolfini'
__title__ = 'Check Nomenclatura Griglie e Livelli'

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
parent_dir = os.path.abspath(os.path.join(script_dir,'..', '..','..','000_Raccolta CSV di controllo','12_CSV_Nomenclatura Livelli.csv'))

#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output()

# CREAZIONE LISTE DI OUTPUT DATA
GRID_LEVELS_NAMING_CSV_OUTPUT = []
GRID_LEVELS_NAMING_CSV_OUTPUT.append(["Categoria","Nome Elemento","ID Elemento","Stato"])
##############################################################



def FormattaNumero(numero):
    if numero >= 0:
        Modif = "+{:.2f}m".format(numero)  # Aggiunge '+' per i numeri positivi
    else:
        Modif = "{:.2f}m".format(numero)   # I numeri negativi mantengono il segno
    return Modif




DizionarioDiVerifica = {}

# ACCEDO AL CSV DI CONTROLLO E CREO IL DIZIONARIO DI VERIFICA
with codecs.open(parent_dir, 'r', 'utf-8-sig') as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        Chiave=row[0]
        Codifica=row[1].split("~")
        DizionarioDiVerifica[Chiave] = Codifica

# COLLECTOR GRIGLIE E LIVELLI
Collector_Griglie = FilteredElementCollector(doc).OfClass(Grid).ToElements()
Collector_Livelli = FilteredElementCollector(doc).OfClass(Level).ToElements()

output.print_md("# Verifica Nomenclatura Griglie e Livelli")
output.print_md("---")


# VERIFICA NOME GRIGLIA
DataTable = []
for Griglia in Collector_Griglie:
    if len(Griglia.Name) == 1 and Griglia.Name.isupper() and Griglia.Name.isalpha() or len(Griglia.Name)== 2 and Griglia.Name.isdigit():
            DataTable.append([Griglia.Name,output.linkify((Griglia.Id)),":white_heavy_check_mark:"])
            GRID_LEVELS_NAMING_CSV_OUTPUT.append([Griglia.Category.Name,Griglia.Name,Griglia.Id,1])
    else:
        DataTable.append([Griglia.Name,output.linkify((Griglia.Id)),":cross_mark:"])
        GRID_LEVELS_NAMING_CSV_OUTPUT.append([Griglia.Category.Name,Griglia.Name,Griglia.Id,0])

output.print_table(table_data = DataTable,title = "Verifica Nomenclatura Griglie", columns = ["Categoria","ID Elemento","Verifica"],formats = ["","",""])


VERIFICA = "NUMERO CAMPI ERRATO"
SIMBOLO = ":cross_mark:"

# VERIFICA NOME LIVELLO
DataTable = []
for Livello in Collector_Livelli:
    Nome = Livello.Name
    Elevazione = round(Livello.Elevation * 0.3048, 2)
    VALUE = X
    VERIFICA = None  # Initialize VERIFICA to track error messages

    if len(Nome.split("_")) != 5:
        VERIFICA = "Lunghezza nome livello errata"
        SIMBOLO = ":cross_mark:"
        VALUE = 0
    else:
        SuddivisioneNome = Nome.split("_")
        if SuddivisioneNome[0] not in DizionarioDiVerifica["Disciplina"]:
            VERIFICA = "Codice disciplina errato"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif SuddivisioneNome[1] != "LIV":
            VERIFICA = "Codice categoria errato"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif SuddivisioneNome[2] not in DizionarioDiVerifica["Fase"]:
            VERIFICA = "Codice fase errato / Non presente in Database"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif SuddivisioneNome[3] not in DizionarioDiVerifica["Piano"]:
            VERIFICA = "Codice piano errato / Non presente in Database"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif SuddivisioneNome[4] != FormattaNumero(Elevazione):
            VERIFICA = "Codice livello errato"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif Elevazione < 0 and "-" not in SuddivisioneNome[3]:
            VERIFICA = "Livello negativo senza trattino"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        elif Elevazione > 0 and "-" in SuddivisioneNome[3]:
            VERIFICA = "Livello positivo con trattino"
            SIMBOLO = ":cross_mark:"
            VALUE = 0
        else:
            VERIFICA = "Codifica Corretta"
            SIMBOLO = ":white_heavy_check_mark:"
            VALUE = 1

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
            pass
            """ IN ATTESA DI CONFERME 
            GRID_LEVELS_NAMING_CSV_OUTPUT = []
            GRID_LEVELS_NAMING_CSV_OUTPUT.append(["Nome Verifica","Stato"])
            GRID_LEVELS_NAMING_CSV_OUTPUT.append(["Naming Convention - Nomenclatura Griglie e Livelli.",1])
            gridsandlevels_csv_path = os.path.join(folder, "12_GriglieELivelli_Data.csv")
            with codecs.open(gridsandlevels_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(GRID_LEVELS_NAMING_CSV_OUTPUT)
            """
        else:
            gridsandlevels_csv_path = os.path.join(folder, "12_GriglieELivelli_Data.csv")
            with codecs.open(gridsandlevels_csv_path, mode='w', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(GRID_LEVELS_NAMING_CSV_OUTPUT)
